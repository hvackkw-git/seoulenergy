(function () {
  "use strict";

  const app = document.getElementById("app");
  if (!app) return;

  /** 웹에서 입력하는 변수 (나중 계산에 사용) */
  const inputs = {
    areaM2: "",
    householdCount: "",
    avgDedicatedAreaM2: "85",
    checkGeothermal: false,
    checkPv: false,
    checkBapv: false,
    checkBipv: false,
    geothermalCapacityKw: "",
    solarPvKwp: "",
    solarBapvKwp: "",
    solarBipvKwp: "",
    pipeCheck100A: false,
    pipeCheck150A: false,
    pipeCheck200A: false,
    pipeCheck250A: false,
    pipeCheck300A: false,
    pipeCheck350A: false,
    pipeCheck400A: false,
    pipeLength100A: "",
    pipeLength150A: "",
    pipeLength200A: "",
    pipeLength250A: "",
    pipeLength300A: "",
    pipeLength350A: "",
    pipeLength400A: ""
  };

  inputs.smpWonPerKwh = "100";
  inputs.gasWonPerMj = "22";
  inputs.geothermalOptimize = false;
  inputs.optimizerTargetSelfSufficiency = "";
  inputs.optimizerMode = "investment";

  var COST_PER_M2 = 2500000;

  var currentTab = 1;

  var SAVE_SLOT_KEY = "seoulEnergy_slot_";

  /** 개발자용 디폴트 데이터 (코드에 정의된 프리셋). 있으면 해당 슬롯 버튼 초록색 표시 */
  var DEV_SLOTS = {
    1: {
      areaM2: "220722",
      householdCount: "2235",
      avgDedicatedAreaM2: "72",
      checkGeothermal: true,
      checkPv: true,
      checkBapv: false,
      checkBipv: false,
      geothermalCapacityKw: "3140",
      solarPvKwp: "1712",
      solarBapvKwp: "",
      solarBipvKwp: "",
      pipeCheck100A: false,
      pipeCheck150A: false,
      pipeCheck200A: false,
      pipeCheck250A: false,
      pipeCheck300A: false,
      pipeCheck350A: false,
      pipeCheck400A: false,
      pipeLength100A: "",
      pipeLength150A: "",
      pipeLength200A: "",
      pipeLength250A: "",
      pipeLength300A: "",
      pipeLength350A: "",
      pipeLength400A: "",
      smpWonPerKwh: "100",
      gasWonPerMj: "22",
      geothermalOptimize: true
    },
    2: null,
    3: null,
    4: null,
    5: null,
    6: null,
    7: null,
    8: null,
    9: null,
    10: null
  };

  /** 저장/불러오기 후 전체 갱신 */
  function refreshAll() {
    updateAreaPyeong();
    updateConstructionCost();
    updateGeothermalHouseholds();
    updatePrimaryEnergyRequired();
    updatePrimaryEnergyProduction();
    updateSelfSufficiencyRate();
    updateAcquisitionTaxSave();
    updateAllCosts();
    updateGasWonPerKwh();
    updateGasAndElectricityUsage();
    updateEnergyCosts();
    updateNpvIrrReport();
    updateSensitivityTable();
    updateFinalReport();
  }

  /** 현재 폼 값으로 저장용 객체 생성 */
  function getInputsForSave() {
    var data = {};
    var key;
    for (key in inputs) {
      if (!inputs.hasOwnProperty(key)) continue;
      var el = document.getElementById(key);
      if (el) {
        if (el.type === "checkbox") data[key] = el.checked;
        else data[key] = el.value;
      } else {
        data[key] = inputs[key];
      }
    }
    return data;
  }

  /** 불러온 데이터를 inputs와 DOM에 반영 후 갱신 */
  function applyLoadedData(data) {
    var key;
    for (key in data) {
      if (!data.hasOwnProperty(key)) continue;
      inputs[key] = data[key];
      var el = document.getElementById(key);
      if (el) {
        if (el.type === "checkbox") el.checked = !!data[key];
        else el.value = data[key];
      }
    }
    [["checkGeothermal", "renewableRowGeothermal"], ["checkPv", "renewableRowPv"], ["checkBapv", "renewableRowBapv"], ["checkBipv", "renewableRowBipv"]].forEach(function (arr) {
      var row = document.getElementById(arr[1]);
      if (row) row.style.display = inputs[arr[0]] ? "block" : "none";
    });
    ["100A", "150A", "200A", "250A", "300A", "350A", "400A"].forEach(function (suffix) {
      var row = document.getElementById("pipeRow" + suffix);
      if (row) row.style.display = inputs["pipeCheck" + suffix] ? "block" : "none";
    });
    var mode = inputs.optimizerMode;
    if (mode === "investment" || mode === "economics") {
      app.querySelectorAll(".optimizer-toggle-btn").forEach(function (b) {
        b.classList.toggle("is-active", b.getAttribute("data-optimizer-mode") === mode);
      });
    }
    refreshAll();
  }

  function updateSlotButtons() {
    for (var n = 1; n <= 10; n++) {
      var btn = document.querySelector(".slot-btn[data-slot=\"" + n + "\"]");
      if (!btn) continue;
      var raw = typeof localStorage !== "undefined" ? localStorage.getItem(SAVE_SLOT_KEY + n) : null;
      if (raw) btn.classList.add("is-saved"); else btn.classList.remove("is-saved");
    }
  }

  function updateDevSlotButtons() {
    for (var n = 1; n <= 10; n++) {
      var btn = document.querySelector(".dev-slot-btn[data-dev-slot=\"" + n + "\"]");
      if (!btn) continue;
      if (DEV_SLOTS[n]) btn.classList.add("is-saved"); else btn.classList.remove("is-saved");
    }
  }

  /** 슬롯 버튼·저장 버튼 바인딩 (render 내부에서 호출) */
  function bindSlotAndSaveButtons() {
    app.querySelectorAll(".dev-slot-btn").forEach(function (btn) {
      btn.addEventListener("click", function () {
        var n = parseInt(btn.getAttribute("data-dev-slot"), 10);
        if (!Number.isFinite(n) || n < 1 || n > 10) return;
        var data = DEV_SLOTS[n];
        if (!data || typeof data !== "object") return;
        applyLoadedData(data);
      });
    });
    app.querySelectorAll(".slot-btn").forEach(function (btn) {
      btn.addEventListener("click", function () {
        var n = parseInt(btn.getAttribute("data-slot"), 10);
        if (!Number.isFinite(n) || n < 1 || n > 10) return;
        var raw = typeof localStorage !== "undefined" ? localStorage.getItem(SAVE_SLOT_KEY + n) : null;
        if (!raw) return;
        try {
          var data = JSON.parse(raw);
          if (data && typeof data === "object") applyLoadedData(data);
        } catch (e) {}
      });
    });
    [["saveBtnTab1", "saveSlotTab1"]].forEach(function (pair) {
      var saveBtn = document.getElementById(pair[0]);
      var slotSelect = document.getElementById(pair[1]);
      if (saveBtn && slotSelect) {
        saveBtn.addEventListener("click", function () {
          var n = parseInt(slotSelect.value, 10);
          if (!Number.isFinite(n) || n < 1 || n > 10) return;
          var data = getInputsForSave();
          try {
            if (typeof localStorage !== "undefined") {
              localStorage.setItem(SAVE_SLOT_KEY + n, JSON.stringify(data));
              updateSlotButtons();
            }
          } catch (e) {}
        });
      }
    });
    var deleteBtn = document.getElementById("saveDeleteBtn1");
    var slotSelectForDelete = document.getElementById("saveSlotTab1");
    if (deleteBtn && slotSelectForDelete) {
      deleteBtn.addEventListener("click", function () {
        var n = parseInt(slotSelectForDelete.value, 10);
        if (!Number.isFinite(n) || n < 1 || n > 10) return;
        try {
          if (typeof localStorage !== "undefined") {
            localStorage.removeItem(SAVE_SLOT_KEY + n);
            updateSlotButtons();
          }
        } catch (e) {}
      });
    }
    updateSlotButtons();
    updateDevSlotButtons();
  }

  /** ? 버튼용 도움말 메모 (data-help id → 메모 텍스트, 단가는 getHelpText에서 처리) */
  var HELP_MEMOS = {
    areaM2: "공사비 기준: '25년 건축정보시스템 평균 주거용 건축 공사비",
    householdCount: "지열 3USRT = 85㎡ 1세대 공급",
    avgDedicatedAreaM2: "국민평형 = 85㎡",
    pipeCheck: "열수송관 구경별 m당 단가는 각 행의 ? 를 누르면 확인할 수 있습니다.",
    reportPrimaryEnergyRequired: "연면적 × 연간 단위면적당 에너지소요량(120 kWh/㎡·yr)으로 계산합니다. 자립률 계산의 분모로 사용됩니다.",
    reportPrimaryEnergyProduction: "지열·PV·BAPV·BIPV 설비가 연간 생산하는 1차에너지 합계입니다. 자립률 계산의 분자로 사용됩니다.",
    reportSelfSufficiencyRate: "1차에너지생산량 ÷ 1차에너지소요량 × 100 입니다. 이 값으로 ZEB 등급(1~5등급)을 판정합니다.",
    reportAcquisitionTaxSave: "공사비의 3.16%에 ZEB 등급별 비율(1~3등급 20%, 4등급 18%, 5등급 15%)을 곱한 취득세 절감액입니다.",
    reportTotalEnergyFacilityCost: "지열·PV·BAPV·BIPV·열수송관 비용을 모두 합산한 추가 에너지설비 투자비입니다.",
    reportCostAfterAcquisitionTaxSave: "추가 에너지설비 투자비에서 취득세 절감액을 뺀, 취득세 경감을 반영한 실질 투자비입니다.",
    reportAdditionalEnergyFacilityCostTab2: "경제성 분석·간이 투자회수기간에 사용하는 추가 에너지설비 투자비입니다. 지열 설치로 보일러·에어컨을 설치하지 않아도 되는 효과를 반영하여, 지열 투자비의 3/4만 반영합니다(대체 설비 비용을 1/4로 가정). '지열 최적화' 체크 시 지열 비용을 절반으로 보는 옵션도 함께 적용됩니다.",
    reportAcquisitionTaxSaveTab2: "공사비의 3.16%에 ZEB 등급별 비율(1~3등급 20%, 4등급 18%, 5등급 15%)을 곱한 취득세 절감액입니다.",
    economicsElectricityUse: "연면적·전력 비중(37.5%)·120 kWh/㎡·yr 등으로 산출한 연간 전력 사용량입니다.",
    economicsElectricityPrice: "전기 단가(원/kWh). SMP 등 적용 단가를 입력합니다.",
    economicsGasUse: "연면적·가스 비중(62.5%)·120 kWh/㎡·yr 등으로 산출한 연간 가스 사용량입니다.",
    economicsGasPrice: "가스 단가(원/MJ).",
    economicsElectricitySave: "지열 냉방 대체(COP 기준) 절감과 PV·BAPV·BIPV 발전분(1차에너지÷2.75×단가)을 지열세대 비율로 반영한 전력요금 절감액입니다.",
    economicsGasSave: "지열 난방·급탕 대체(COP 4)로 인한 가스요금 절감에서 지열 대체 전력 비용을 뺀 금액을 지열세대 비율로 반영한 값입니다.",
    economicsTotalSave: "전력요금 절감과 가스요금 절감의 합계입니다.",
    economicsPaybackPeriod: "(추가 에너지설비 투자비 − 취득세 절감) ÷ 총 에너지비용 절감으로 구한 간이 투자회수기간(연)입니다.",
    economicsAreaChangeCost: "면적 증감에 따른 비용입니다."
  };

  function render() {
    app.innerHTML = `
      <div class="card">
        <div class="card-header">
          <h1>SE Platform</h1>
          <div class="tab-buttons">
            <button type="button" class="tab-btn ${currentTab === 1 ? "is-active" : ""}" data-tab="1">투자비산출</button>
            <button type="button" class="tab-btn ${currentTab === 2 ? "is-active" : ""}" data-tab="2">경제성분석</button>
            <button type="button" class="tab-btn ${currentTab === 3 ? "is-active" : ""}" data-tab="3">민감도분석</button>
            <button type="button" class="tab-btn ${currentTab === 4 ? "is-active" : ""}" data-tab="4">최적값찾기</button>
            <button type="button" class="tab-btn ${currentTab === 5 ? "is-active" : ""}" data-tab="5">최종보고서</button>
            <button type="button" class="tab-btn ${currentTab === 6 ? "is-active" : ""}" data-tab="6">간이계산기</button>
          </div>
          <div class="save-slot-frame">
            <div class="developer-slot-buttons" aria-label="개발자 프리셋">
              ${[1,2,3,4,5,6,7,8,9,10].map(function(n) { return "<button type=\"button\" class=\"dev-slot-btn\" data-dev-slot=\"" + n + "\" aria-label=\"개발자 슬롯 " + n + "\">" + n + "</button>"; }).join("")}
            </div>
            <div class="save-slot-second-row">
              <div class="slot-buttons" aria-label="저장 슬롯 불러오기">
                ${[1,2,3,4,5,6,7,8,9,10].map(function(n) { return "<button type=\"button\" class=\"slot-btn\" data-slot=\"" + n + "\" aria-label=\"슬롯 " + n + " 불러오기\">" + n + "</button>"; }).join("")}
              </div>
              <div class="save-row">
                <label class="save-row-label">저장 슬롯</label>
                <select id="saveSlotTab1" class="save-slot-select" aria-label="저장 슬롯">
                  ${[1,2,3,4,5,6,7,8,9,10].map(function(n) { return "<option value=\"" + n + "\">" + n + "</option>"; }).join("")}
                </select>
                <button type="button" id="saveBtnTab1" class="save-submit-btn">저장</button>
                <button type="button" id="saveDeleteBtn1" class="save-delete-btn">삭제</button>
              </div>
            </div>
          </div>
        </div>
        <div id="tabPanel1" class="tab-panel ${currentTab === 1 ? "is-active" : ""}">

        <div class="form-two-cols form-input-cols">
          <div class="form-col form-col-left">
            <div class="form-section form-section-grid">
              <div class="form-row">
                <div class="form-row-head">
                  <label>연면적 <span class="unit">(m²)</span></label>
                  <button type="button" class="help-btn" aria-label="도움말" data-help="areaM2">?</button>
                </div>
                <input type="number" id="areaM2" placeholder="예: 10000" min="0" step="0.01" value="${inputs.areaM2}" />
                <span class="result-frame result-frame-construction" id="constructionCost">공사비: —</span>
                <span class="result-frame result-frame-pyeong"><span id="areaPyeong">— 평</span></span>
              </div>
              <div class="form-row">
                <div class="form-row-head">
                  <label>세대수 <span class="unit">(세대)</span></label>
                  <button type="button" class="help-btn" aria-label="도움말" data-help="householdCount">?</button>
                </div>
                <input type="number" id="householdCount" placeholder="예: 100" min="0" step="1" value="${inputs.householdCount}" />
                <span class="result-frame result-frame-household">
                  <span id="geothermalHouseholds">지열세대: —</span><br/>
                  <span id="nonGeothermalHouseholds">비 지열세대: —</span>
                </span>
              </div>
              <div class="form-row">
                <div class="form-row-head">
                  <label>평균 전용면적 <span class="unit">(m²)</span></label>
                  <button type="button" class="help-btn" aria-label="도움말" data-help="avgDedicatedAreaM2">?</button>
                </div>
                <input type="number" id="avgDedicatedAreaM2" placeholder="예: 85" min="0" step="0.01" value="${inputs.avgDedicatedAreaM2}" />
              </div>
            </div>
          </div>
          <div class="form-col form-col-middle">
            <div class="form-section form-section-grid">
              <div class="form-row renewable-check-row">
                <div class="form-row-head">
                  <span class="form-row-head-label">신재생 설비</span>
                </div>
                <div class="pipe-check-wrap">
                  <label class="pipe-check"><input type="checkbox" id="checkGeothermal" ${inputs.checkGeothermal ? "checked" : ""}> 지열</label>
                  <label class="pipe-check"><input type="checkbox" id="checkPv" ${inputs.checkPv ? "checked" : ""}> PV</label>
                  <label class="pipe-check"><input type="checkbox" id="checkBapv" ${inputs.checkBapv ? "checked" : ""}> BAPV</label>
                  <label class="pipe-check"><input type="checkbox" id="checkBipv" ${inputs.checkBipv ? "checked" : ""}> BIPV</label>
                </div>
              </div>
              <div id="renewableRowGeothermal" class="form-row renewable-input-row" style="display:${inputs.checkGeothermal ? "block" : "none"}">
                <div class="form-row-head">
                  <label>지열 설치용량 <span class="unit">(kW)</span></label>
                  <button type="button" class="help-btn" aria-label="도움말" data-help="geothermalCapacityKw">?</button>
                </div>
                <input type="number" id="geothermalCapacityKw" placeholder="예: 50" min="0" step="0.01" value="${inputs.geothermalCapacityKw}" />
                <span id="costGeothermal" class="cost-display">—</span>
              </div>
              <div id="renewableRowPv" class="form-row renewable-input-row" style="display:${inputs.checkPv ? "block" : "none"}">
                <div class="form-row-head">
                  <label>PV 설치용량 <span class="unit">(kWp)</span></label>
                  <button type="button" class="help-btn" aria-label="도움말" data-help="solarPvKwp">?</button>
                </div>
                <input type="number" id="solarPvKwp" placeholder="예: 0" min="0" step="0.01" value="${inputs.solarPvKwp}" />
                <span id="costPv" class="cost-display">—</span>
              </div>
              <div id="renewableRowBapv" class="form-row renewable-input-row" style="display:${inputs.checkBapv ? "block" : "none"}">
                <div class="form-row-head">
                  <label>BAPV 설치용량 <span class="unit">(kWp)</span></label>
                  <button type="button" class="help-btn" aria-label="도움말" data-help="solarBapvKwp">?</button>
                </div>
                <input type="number" id="solarBapvKwp" placeholder="예: 0" min="0" step="0.01" value="${inputs.solarBapvKwp}" />
                <span id="costBapv" class="cost-display">—</span>
              </div>
              <div id="renewableRowBipv" class="form-row renewable-input-row" style="display:${inputs.checkBipv ? "block" : "none"}">
                <div class="form-row-head">
                  <label>BIPV 설치용량 <span class="unit">(kWp)</span></label>
                  <button type="button" class="help-btn" aria-label="도움말" data-help="solarBipvKwp">?</button>
                </div>
                <input type="number" id="solarBipvKwp" placeholder="예: 0" min="0" step="0.01" value="${inputs.solarBipvKwp}" />
                <span id="costBipv" class="cost-display">—</span>
              </div>
              <div class="form-row pipe-check-row">
                <div class="form-row-head">
                  <span class="form-row-head-label">열수송관 구경</span>
                  <button type="button" class="help-btn" aria-label="도움말" data-help="pipeCheck">?</button>
                </div>
                <div class="pipe-check-wrap">
                  <label class="pipe-check"><input type="checkbox" id="pipeCheck100A" ${inputs.pipeCheck100A ? "checked" : ""}> 100A</label>
                  <label class="pipe-check"><input type="checkbox" id="pipeCheck150A" ${inputs.pipeCheck150A ? "checked" : ""}> 150A</label>
                  <label class="pipe-check"><input type="checkbox" id="pipeCheck200A" ${inputs.pipeCheck200A ? "checked" : ""}> 200A</label>
                  <label class="pipe-check"><input type="checkbox" id="pipeCheck250A" ${inputs.pipeCheck250A ? "checked" : ""}> 250A</label>
                  <label class="pipe-check"><input type="checkbox" id="pipeCheck300A" ${inputs.pipeCheck300A ? "checked" : ""}> 300A</label>
                  <label class="pipe-check"><input type="checkbox" id="pipeCheck350A" ${inputs.pipeCheck350A ? "checked" : ""}> 350A</label>
                  <label class="pipe-check"><input type="checkbox" id="pipeCheck400A" ${inputs.pipeCheck400A ? "checked" : ""}> 400A</label>
                </div>
              </div>
              <div id="pipeRow100A" class="form-row pipe-input-row" style="display:${inputs.pipeCheck100A ? "block" : "none"}">
                <div class="form-row-head">
                  <label>열수송관 길이 100A <span class="unit">(m)</span></label>
                  <button type="button" class="help-btn" aria-label="도움말" data-help="pipeLength100A">?</button>
                </div>
                <input type="number" id="pipeLength100A" placeholder="예: 0" min="0" step="0.01" value="${inputs.pipeLength100A}" />
                <span id="costPipe100A" class="cost-display">—</span>
              </div>
              <div id="pipeRow150A" class="form-row pipe-input-row" style="display:${inputs.pipeCheck150A ? "block" : "none"}">
                <div class="form-row-head">
                  <label>열수송관 길이 150A <span class="unit">(m)</span></label>
                  <button type="button" class="help-btn" aria-label="도움말" data-help="pipeLength150A">?</button>
                </div>
                <input type="number" id="pipeLength150A" placeholder="예: 0" min="0" step="0.01" value="${inputs.pipeLength150A}" />
                <span id="costPipe150A" class="cost-display">—</span>
              </div>
              <div id="pipeRow200A" class="form-row pipe-input-row" style="display:${inputs.pipeCheck200A ? "block" : "none"}">
                <div class="form-row-head">
                  <label>열수송관 길이 200A <span class="unit">(m)</span></label>
                  <button type="button" class="help-btn" aria-label="도움말" data-help="pipeLength200A">?</button>
                </div>
                <input type="number" id="pipeLength200A" placeholder="예: 0" min="0" step="0.01" value="${inputs.pipeLength200A}" />
                <span id="costPipe200A" class="cost-display">—</span>
              </div>
              <div id="pipeRow250A" class="form-row pipe-input-row" style="display:${inputs.pipeCheck250A ? "block" : "none"}">
                <div class="form-row-head">
                  <label>열수송관 길이 250A <span class="unit">(m)</span></label>
                  <button type="button" class="help-btn" aria-label="도움말" data-help="pipeLength250A">?</button>
                </div>
                <input type="number" id="pipeLength250A" placeholder="예: 0" min="0" step="0.01" value="${inputs.pipeLength250A}" />
                <span id="costPipe250A" class="cost-display">—</span>
              </div>
              <div id="pipeRow300A" class="form-row pipe-input-row" style="display:${inputs.pipeCheck300A ? "block" : "none"}">
                <div class="form-row-head">
                  <label>열수송관 길이 300A <span class="unit">(m)</span></label>
                  <button type="button" class="help-btn" aria-label="도움말" data-help="pipeLength300A">?</button>
                </div>
                <input type="number" id="pipeLength300A" placeholder="예: 0" min="0" step="0.01" value="${inputs.pipeLength300A}" />
                <span id="costPipe300A" class="cost-display">—</span>
              </div>
              <div id="pipeRow350A" class="form-row pipe-input-row" style="display:${inputs.pipeCheck350A ? "block" : "none"}">
                <div class="form-row-head">
                  <label>열수송관 길이 350A <span class="unit">(m)</span></label>
                  <button type="button" class="help-btn" aria-label="도움말" data-help="pipeLength350A">?</button>
                </div>
                <input type="number" id="pipeLength350A" placeholder="예: 0" min="0" step="0.01" value="${inputs.pipeLength350A}" />
                <span id="costPipe350A" class="cost-display">—</span>
              </div>
              <div id="pipeRow400A" class="form-row pipe-input-row" style="display:${inputs.pipeCheck400A ? "block" : "none"}">
                <div class="form-row-head">
                  <label>열수송관 길이 400A <span class="unit">(m)</span></label>
                  <button type="button" class="help-btn" aria-label="도움말" data-help="pipeLength400A">?</button>
                </div>
                <input type="number" id="pipeLength400A" placeholder="예: 0" min="0" step="0.01" value="${inputs.pipeLength400A}" />
                <span id="costPipe400A" class="cost-display">—</span>
              </div>
            </div>
          </div>
        </div>

        <div class="report-tab">
          <h2 class="report-tab-title">보고서</h2>
          <div class="report-frame form-three-cols">
            <div class="report-col report-col-1">
              <div class="report-row">
                <div class="report-row-head">
                  <label>1차에너지 소요량 <span class="unit">(자립률 분모)</span></label>
                  <button type="button" class="help-btn" aria-label="도움말" data-help="reportPrimaryEnergyRequired">?</button>
                </div>
                <span id="primaryEnergyRequired" class="report-value">—</span>
              </div>
            </div>
            <div class="report-col report-col-2">
              <div class="report-row">
                <div class="report-row-head">
                  <label>1차에너지생산량 <span class="unit">(자립률 분자)</span></label>
                  <button type="button" class="help-btn" aria-label="도움말" data-help="reportPrimaryEnergyProduction">?</button>
                </div>
                <span id="primaryEnergyProduction" class="report-value">—</span>
              </div>
              <div class="report-row">
                <div class="report-row-head">
                  <label>자립률 <span class="unit">(생산량/소요량×100)</span></label>
                  <button type="button" class="help-btn" aria-label="도움말" data-help="reportSelfSufficiencyRate">?</button>
                </div>
                <span id="selfSufficiencyRate" class="report-value">—</span>
                <span id="reportGradeLabel" class="report-grade">—</span>
              </div>
              <div class="report-row">
                <div class="report-row-head">
                  <label>취득세 절감 <span class="unit">(ZEB 등급 적용)</span></label>
                  <button type="button" class="help-btn" aria-label="도움말" data-help="reportAcquisitionTaxSave">?</button>
                </div>
                <span id="acquisitionTaxSave" class="report-value">—</span>
              </div>
            </div>
            <div class="report-col report-col-3">
              <div class="report-row">
                <div class="report-row-head">
                  <label>추가 에너지설비 투자비</label>
                  <button type="button" class="help-btn" aria-label="도움말" data-help="reportTotalEnergyFacilityCost">?</button>
                </div>
                <span id="totalEnergyFacilityCost" class="report-value report-value-total">—</span>
              </div>
              <div class="report-row">
                <div class="report-row-head">
                  <label>취득세 경감 적용 투자비</label>
                  <button type="button" class="help-btn" aria-label="도움말" data-help="reportCostAfterAcquisitionTaxSave">?</button>
                </div>
                <span id="costAfterAcquisitionTaxSave" class="report-value report-value-total">—</span>
              </div>
            </div>
          </div>
        </div>
        </div>

        <div id="tabPanel2" class="tab-panel ${currentTab === 2 ? "is-active" : ""}">
          <div class="economics-cols form-two-cols form-input-cols">
            <div class="form-col form-col-left economics-col-left">
              <div class="economics-section">
                <div class="economics-row">
                  <div class="economics-row-head">
                    <label>추가 에너지설비 투자비</label>
                    <button type="button" class="help-btn" aria-label="도움말" data-help="reportAdditionalEnergyFacilityCostTab2">?</button>
                  </div>
                  <div class="economics-row-value-wrap">
                    <span id="totalEnergyFacilityCostTab2" class="economics-value">—</span>
                    <label class="economics-check"><input type="checkbox" id="geothermalOptimize" ${inputs.geothermalOptimize ? "checked" : ""}> 지열 최적화</label>
                  </div>
                </div>
                <div class="economics-row">
                  <div class="economics-row-head">
                    <label>취득세 절감</label>
                    <button type="button" class="help-btn" aria-label="도움말" data-help="reportAcquisitionTaxSaveTab2">?</button>
                  </div>
                  <span id="acquisitionTaxSaveTab2" class="economics-value">—</span>
                </div>
                <div class="economics-row">
                  <div class="economics-row-head">
                    <label>면적 증감 비용</label>
                    <button type="button" class="help-btn" aria-label="도움말" data-help="economicsAreaChangeCost">?</button>
                  </div>
                  <span id="areaChangeCostTab2" class="economics-value">—</span>
                </div>
              </div>
            </div>
            <div class="form-col form-col-middle economics-col-right">
              <div class="economics-section">
                <div class="economics-row">
                  <div class="economics-row-head">
                    <label>전기사용량 <span class="unit">(자동 산출)</span></label>
                    <button type="button" class="help-btn" aria-label="도움말" data-help="economicsElectricityUse">?</button>
                  </div>
                  <span id="totalElectricityUseKwh" class="economics-value">—</span>
                </div>
                <div class="economics-row">
                  <div class="economics-row-head">
                    <label>전기요금 <span class="unit">(원/kWh)</span></label>
                    <button type="button" class="help-btn" aria-label="도움말" data-help="economicsElectricityPrice">?</button>
                  </div>
                  <input type="number" id="smpWonPerKwh" placeholder="100" min="0" step="0.01" value="${inputs.smpWonPerKwh}" />
                </div>
                <div class="economics-row">
                  <div class="economics-row-head">
                    <label>가스사용량 <span class="unit">(자동 산출)</span></label>
                    <button type="button" class="help-btn" aria-label="도움말" data-help="economicsGasUse">?</button>
                  </div>
                  <span id="totalGasUseKwh" class="economics-value">—</span>
                  <span id="totalGasUseMj" class="economics-value-sub">—</span>
                </div>
                <div class="economics-row">
                  <div class="economics-row-head">
                    <label>가스요금 <span class="unit">(원/MJ)</span></label>
                    <button type="button" class="help-btn" aria-label="도움말" data-help="economicsGasPrice">?</button>
                  </div>
                  <input type="number" id="gasWonPerMj" placeholder="22" min="0" step="0.01" value="${inputs.gasWonPerMj}" />
                  <span id="gasWonPerKwh" class="economics-calc">— 원/kWh</span>
                </div>
              </div>
            </div>
          </div>
          <div class="economics-report-tab">
            <h2 class="report-tab-title">최종 레포트</h2>
            <div class="report-frame economics-report-frame">
              <div class="report-col report-col-1">
                <div class="report-row">
                  <div class="report-row-head">
                    <label>전력요금 절감</label>
                    <button type="button" class="help-btn" aria-label="도움말" data-help="economicsElectricitySave">?</button>
                  </div>
                  <span id="electricityCostSave" class="report-value">—</span>
                </div>
              </div>
              <div class="report-col report-col-2">
                <div class="report-row">
                  <div class="report-row-head">
                    <label>가스요금 절감</label>
                    <button type="button" class="help-btn" aria-label="도움말" data-help="economicsGasSave">?</button>
                  </div>
                  <span id="gasCostSave" class="report-value">—</span>
                </div>
              </div>
              <div class="report-col report-col-3">
                <div class="report-row">
                  <div class="report-row-head">
                    <label>총 에너지비용 절감</label>
                    <button type="button" class="help-btn" aria-label="도움말" data-help="economicsTotalSave">?</button>
                  </div>
                  <span id="totalEnergyCostSave" class="report-value report-value-total">—</span>
                </div>
                <div class="report-row">
                  <div class="report-row-head">
                    <label>간이 투자회수기간</label>
                    <button type="button" class="help-btn" aria-label="도움말" data-help="economicsPaybackPeriod">?</button>
                  </div>
                  <span id="paybackPeriod" class="report-value">—</span>
                </div>
              </div>
            </div>
            <div class="npv-irr-section">
              <p class="npv-irr-assumptions" id="npvIrrAssumptions">자기자본 100% · 할인율 4.5% · 물가상승 2% · 20년 · 세율 25%</p>
              <div class="npv-irr-result">
                <div class="npv-irr-item">
                  <span class="npv-irr-label">NPV</span>
                  <span id="npvValue" class="npv-irr-value">—</span>
                </div>
                <div class="npv-irr-item">
                  <span class="npv-irr-label">IRR</span>
                  <span id="irrValue" class="npv-irr-value">—</span>
                </div>
              </div>
              <div class="projection-table-wrap">
                <table class="projection-table" aria-label="20년 세전·세후이익 프로젝션">
                  <thead>
                    <tr>
                      <th>연도</th>
                      <th>세전이익</th>
                      <th>세금</th>
                      <th>세후이익</th>
                      <th>NPV</th>
                    </tr>
                  </thead>
                  <tbody id="projectionTableBody">
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        </div>
        </div>

        <div id="tabPanel3" class="tab-panel ${currentTab === 3 ? "is-active" : ""}">
          <div class="sensitivity-section">
            <div class="sensitivity-controls">
              <div class="sensitivity-control-row">
                <label for="sensitivityRange">민감도 범위 <span class="unit">(%)</span></label>
                <input type="number" id="sensitivityRange" value="10" min="1" max="50" step="1" />
              </div>
              <div class="sensitivity-control-row">
                <label for="sensitivityVariable">변수 선택</label>
                <select id="sensitivityVariable" aria-label="민감도 변수">
                  <option value="investment">투자비</option>
                  <option value="smp">전기세</option>
                  <option value="gas">가스요금</option>
                  <option value="taxSave">취득세절감</option>
                </select>
              </div>
            </div>
            <div class="sensitivity-table-wrap">
              <table class="sensitivity-table" id="sensitivityTable">
                <thead>
                  <tr>
                    <th>지표</th>
                    <th id="sensitivityCol0">—</th>
                    <th id="sensitivityCol1">—</th>
                    <th id="sensitivityCol2">—</th>
                    <th id="sensitivityCol3">—</th>
                    <th id="sensitivityCol4">—</th>
                  </tr>
                </thead>
                <tbody>
                  <tr><td>전력요금 절감</td><td id="sensElec0">—</td><td id="sensElec1">—</td><td id="sensElec2">—</td><td id="sensElec3">—</td><td id="sensElec4">—</td></tr>
                  <tr><td>가스요금 절감</td><td id="sensGas0">—</td><td id="sensGas1">—</td><td id="sensGas2">—</td><td id="sensGas3">—</td><td id="sensGas4">—</td></tr>
                  <tr><td>총 에너지비용 절감</td><td id="sensTotal0">—</td><td id="sensTotal1">—</td><td id="sensTotal2">—</td><td id="sensTotal3">—</td><td id="sensTotal4">—</td></tr>
                  <tr><td>간이투자회수 기간</td><td id="sensPayback0">—</td><td id="sensPayback1">—</td><td id="sensPayback2">—</td><td id="sensPayback3">—</td><td id="sensPayback4">—</td></tr>
                  <tr><td>NPV</td><td id="sensNpv0">—</td><td id="sensNpv1">—</td><td id="sensNpv2">—</td><td id="sensNpv3">—</td><td id="sensNpv4">—</td></tr>
                  <tr><td>IRR</td><td id="sensIrr0">—</td><td id="sensIrr1">—</td><td id="sensIrr2">—</td><td id="sensIrr3">—</td><td id="sensIrr4">—</td></tr>
                </tbody>
              </table>
            </div>
            <div class="sensitivity-charts">
              <div class="sensitivity-chart-box">
                <h4 class="sensitivity-chart-title">시나리오별 NPV</h4>
                <div class="sensitivity-chart-svg-wrap" id="sensitivityNpvChart"></div>
              </div>
              <div class="sensitivity-chart-box">
                <h4 class="sensitivity-chart-title">시나리오별 IRR</h4>
                <div class="sensitivity-chart-svg-wrap" id="sensitivityIrrChart"></div>
              </div>
            </div>
          </div>
        </div>

        <div id="tabPanel4" class="tab-panel ${currentTab === 4 ? "is-active" : ""}">
          <p class="sub">최적값찾기</p>
          <div class="optimizer-section">
            <div class="optimizer-inputs-row">
              <div class="optimizer-field">
                <label for="optimizerTargetSelfSufficiency">자립률 목표 <span class="unit">(%)</span></label>
                <input type="number" id="optimizerTargetSelfSufficiency" placeholder="예: 40" min="0" max="100" step="0.1" value="${inputs.optimizerTargetSelfSufficiency}" />
              </div>
              <div class="optimizer-field">
                <label for="optimizerAreaM2">연면적 <span class="unit">(m²)</span></label>
                <input type="number" id="optimizerAreaM2" placeholder="예: 10000" min="0" step="0.01" value="${inputs.areaM2 || ""}" />
              </div>
              <div class="optimizer-field">
                <label for="optimizerHouseholdCount">세대수 <span class="unit">(세대)</span></label>
                <input type="number" id="optimizerHouseholdCount" placeholder="예: 100" min="0" step="1" value="${inputs.householdCount || ""}" />
              </div>
              <div class="optimizer-field">
                <label for="optimizerRoofAreaM2">옥상면적 <span class="unit">(m²)</span></label>
                <input type="number" id="optimizerRoofAreaM2" placeholder="예: 500" min="0" step="0.01" />
              </div>
              <div class="optimizer-field">
                <label for="optimizerSiteAreaM2">부지면적 <span class="unit">(m²)</span></label>
                <input type="number" id="optimizerSiteAreaM2" placeholder="예: 49" min="0" step="0.01" />
              </div>
            </div>
            <div class="optimizer-row">
              <div class="optimizer-row-head">
                <span class="optimizer-row-label">최적화 기준</span>
              </div>
              <div class="optimizer-toggle-wrap">
                <div class="optimizer-toggle" role="group" aria-label="최적화 기준 선택">
                  <button type="button" class="optimizer-toggle-btn ${inputs.optimizerMode === "investment" ? "is-active" : ""}" data-optimizer-mode="investment">초기투자비 중심</button>
                  <button type="button" class="optimizer-toggle-btn ${inputs.optimizerMode === "economics" ? "is-active" : ""}" data-optimizer-mode="economics">경제성 중심</button>
                </div>
                <label class="optimizer-geo-check"><input type="checkbox" id="optimizerGeothermalOptimize" ${inputs.geothermalOptimize ? "checked" : ""}> 지열 최적화</label>
                <button type="button" id="optimizerRunBtn" class="optimizer-run-btn">RUN</button>
              </div>
            </div>
            <div id="optimizerResult" class="optimizer-result" aria-live="polite"></div>
          </div>
        </div>

        <div id="tabPanel5" class="tab-panel ${currentTab === 5 ? "is-active" : ""}">
          <div class="final-report">
            <h2 class="final-report-main-title">최종보고서</h2>
            <p class="final-report-desc">투자비산출·경제성분석·민감도분석 결과를 한눈에 확인합니다.</p>

            <section class="final-report-section">
              <div class="final-report-section-head">
                <h3 class="final-report-section-title">1. 투자비 산출 요약</h3>
                <button type="button" class="final-report-copy-btn" data-section="1">이미지 클립보드로 복사</button>
              </div>
              <div class="final-report-table-wrap">
                <table class="final-report-table">
                  <tbody id="finalReportInvestmentBody"></tbody>
                </table>
              </div>
            </section>

            <section class="final-report-section">
              <div class="final-report-section-head">
                <h3 class="final-report-section-title">2. 경제성 분석 요약</h3>
                <button type="button" class="final-report-copy-btn" data-section="2">이미지 클립보드로 복사</button>
              </div>
              <div class="final-report-table-wrap">
                <table class="final-report-table final-report-table-economics">
                  <tbody id="finalReportEconomicsBody"></tbody>
                </table>
              </div>
            </section>

            <section class="final-report-section">
              <div class="final-report-section-head">
                <h3 class="final-report-section-title">3. 민감도 분석 요약</h3>
                <button type="button" class="final-report-copy-btn" data-section="3">이미지 클립보드로 복사</button>
              </div>
              <div class="final-report-table-wrap">
                <table class="final-report-table final-report-table-sensitivity">
                  <thead>
                    <tr>
                      <th>지표</th>
                      <th id="finalSensCol0">—</th>
                      <th id="finalSensCol1">—</th>
                      <th id="finalSensCol2">—</th>
                      <th id="finalSensCol3">—</th>
                      <th id="finalSensCol4">—</th>
                    </tr>
                  </thead>
                  <tbody id="finalReportSensitivityBody"></tbody>
                </table>
              </div>
              <p class="final-report-note" id="finalReportSensitivityNote">민감도 범위 및 변수는 민감도분석 탭에서 설정한 값을 사용합니다.</p>
              <div class="sensitivity-charts final-sensitivity-charts">
                <div class="sensitivity-chart-box">
                  <h4 class="sensitivity-chart-title">NPV 민감도</h4>
                  <div class="sensitivity-chart-svg-wrap" id="finalSensitivityNpvChart"></div>
                </div>
                <div class="sensitivity-chart-box">
                  <h4 class="sensitivity-chart-title">IRR 민감도</h4>
                  <div class="sensitivity-chart-svg-wrap" id="finalSensitivityIrrChart"></div>
                </div>
              </div>
            </section>
            <div class="final-report-actions">
              <button type="button" id="finalReportPrintBtn" class="final-report-print-btn">출력</button>
              <button type="button" id="finalReportExcelBtn" class="final-report-excel-btn">엑셀</button>
            </div>
          </div>
        </div>

        <div id="tabPanel6" class="tab-panel ${currentTab === 6 ? "is-active" : ""}">
          <p class="sub">간이계산기</p>
          <div class="calc-section">
            <h3 class="calc-title">지열 단위 변환</h3>
            <div class="calc-row">
              <label for="calcGeoKw">지열 용량 <span class="unit">(kW)</span></label>
              <input type="number" id="calcGeoKw" placeholder="예: 100" min="0" step="0.01" />
              <span class="calc-result">→ USRT: <span id="calcGeoUsrt">—</span></span>
            </div>
          </div>
          <div class="calc-section">
            <h3 class="calc-title">지열 개략 투자비</h3>
            <div class="calc-row">
              <label for="calcGeoCostKw">지열 용량 <span class="unit">(kW)</span></label>
              <input type="number" id="calcGeoCostKw" placeholder="예: 50" min="0" step="0.01" />
              <span class="calc-result">→ 투자비: <span id="calcGeoCost">—</span></span>
            </div>
          </div>
          <div class="calc-section">
            <h3 class="calc-title">면적당 태양광 용량</h3>
            <p class="calc-desc">1 m²당 100 W 기준</p>
            <div class="calc-row">
              <label for="calcSolarAreaM2">면적 <span class="unit">(m²)</span></label>
              <input type="number" id="calcSolarAreaM2" placeholder="예: 100" min="0" step="0.01" />
              <span class="calc-result">→ <span id="calcSolarKwp">—</span> kW</span>
            </div>
          </div>
          <div class="calc-section">
            <h3 class="calc-title">옥상 최대 설치 용량</h3>
            <p class="calc-desc">옥상면적 × 70% × 100 W</p>
            <div class="calc-row">
              <label for="calcRoofAreaM2">옥상면적 <span class="unit">(m²)</span></label>
              <input type="number" id="calcRoofAreaM2" placeholder="예: 500" min="0" step="0.01" />
              <span class="calc-result">→ 최대 <span id="calcRoofKwp">—</span> kW</span>
            </div>
          </div>
          <div class="calc-section">
            <h3 class="calc-title">태양광 개략 투자비</h3>
            <div class="calc-row">
              <label for="calcPvCostKwp">태양광 용량 <span class="unit">(kW)</span></label>
              <input type="number" id="calcPvCostKwp" placeholder="예: 100" min="0" step="0.01" />
              <span class="calc-result">→ 투자비: <span id="calcPvCost">—</span></span>
            </div>
          </div>
        </div>
      </div>
    `;

    document.querySelectorAll(".tab-btn").forEach(function (btn) {
      btn.addEventListener("click", function () {
        var tab = parseInt(btn.getAttribute("data-tab"), 10);
        currentTab = tab;
        app.querySelectorAll(".tab-btn").forEach(function (b) { b.classList.remove("is-active"); });
        app.querySelectorAll(".tab-panel").forEach(function (p) { p.classList.remove("is-active"); });
        btn.classList.add("is-active");
        var panel = document.getElementById("tabPanel" + tab);
        if (panel) panel.classList.add("is-active");
        if (tab === 3) updateSensitivityTable();
        if (tab === 4) {
          var oa = document.getElementById("optimizerAreaM2");
          if (oa) oa.value = inputs.areaM2 || "";
          var oh = document.getElementById("optimizerHouseholdCount");
          if (oh) oh.value = inputs.householdCount || "";
          var oGeo = document.getElementById("optimizerGeothermalOptimize");
          if (oGeo) oGeo.checked = !!inputs.geothermalOptimize;
        }
        if (tab === 5) updateFinalReport();
      });
    });

    var tab2Ids = ["smpWonPerKwh", "gasWonPerMj"];
    tab2Ids.forEach(function (id) {
      var el = document.getElementById(id);
      if (el) {
        el.addEventListener("input", function () {
          inputs[id] = el.value;
          updateGasWonPerKwh();
          updateEnergyCosts();
        });
        el.addEventListener("change", function () {
          inputs[id] = el.value;
          updateGasWonPerKwh();
          updateEnergyCosts();
        });
      }
    });

    var geoOptEl = document.getElementById("geothermalOptimize");
    if (geoOptEl) {
      geoOptEl.addEventListener("change", function () {
        inputs.geothermalOptimize = geoOptEl.checked;
        updateTotalEnergyFacilityCost();
      });
    }

    var ids = [
      "areaM2", "householdCount", "avgDedicatedAreaM2",
      "geothermalCapacityKw", "solarPvKwp", "solarBapvKwp", "solarBipvKwp",
      "pipeLength100A", "pipeLength150A", "pipeLength200A", "pipeLength250A",
      "pipeLength300A", "pipeLength350A", "pipeLength400A"
    ];
    ids.forEach(function (id) {
      var el = document.getElementById(id);
      if (el) {
        el.addEventListener("input", function () {
          inputs[id] = el.value;
          updateAreaPyeong();
          updateConstructionCost();
          updateGeothermalHouseholds();
          updatePrimaryEnergyRequired();
          updatePrimaryEnergyProduction();
          updateSelfSufficiencyRate();
          updateAcquisitionTaxSave();
          updateAllCosts();
          updateGasAndElectricityUsage();
          updateEnergyCosts();
        });
        el.addEventListener("change", function () {
          inputs[id] = el.value;
          updateAreaPyeong();
          updateConstructionCost();
          updateGeothermalHouseholds();
          updatePrimaryEnergyRequired();
          updatePrimaryEnergyProduction();
          updateSelfSufficiencyRate();
          updateAcquisitionTaxSave();
          updateAllCosts();
          updateGasAndElectricityUsage();
          updateEnergyCosts();
        });
      }
    });
    ["100A", "150A", "200A", "250A", "300A", "350A", "400A"].forEach(function (suffix) {
      var id = "pipeCheck" + suffix;
      var rowId = "pipeRow" + suffix;
      var lengthId = "pipeLength" + suffix;
      var el = document.getElementById(id);
      if (el) {
        el.addEventListener("change", function () {
          inputs[id] = el.checked;
          if (!el.checked) {
            inputs[lengthId] = "";
            var inputEl = document.getElementById(lengthId);
            if (inputEl) inputEl.value = "";
          }
          var row = document.getElementById(rowId);
          if (row) row.style.display = el.checked ? "block" : "none";
          updateAllCosts();
          updateTotalEnergyFacilityCost();
          updateCostAfterAcquisitionTaxSave();
        });
      }
    });

    [["checkGeothermal", "renewableRowGeothermal", "geothermalCapacityKw"], ["checkPv", "renewableRowPv", "solarPvKwp"], ["checkBapv", "renewableRowBapv", "solarBapvKwp"], ["checkBipv", "renewableRowBipv", "solarBipvKwp"]].forEach(function (arr) {
      var checkId = arr[0];
      var rowId = arr[1];
      var inputId = arr[2];
      var el = document.getElementById(checkId);
      if (el) {
        el.addEventListener("change", function () {
          inputs[checkId] = el.checked;
          if (!el.checked) {
            inputs[inputId] = "";
            var inputEl = document.getElementById(inputId);
            if (inputEl) inputEl.value = "";
          }
          var row = document.getElementById(rowId);
          if (row) row.style.display = el.checked ? "block" : "none";
          updateGeothermalHouseholds();
          updatePrimaryEnergyProduction();
          updateSelfSufficiencyRate();
          updateAcquisitionTaxSave();
          updateAllCosts();
          updateCostAfterAcquisitionTaxSave();
        });
      }
    });

    var popover = document.getElementById("help-popover");
    if (!popover) {
      popover = document.createElement("div");
      popover.id = "help-popover";
      popover.className = "help-popover";
      document.body.appendChild(popover);
    }
    var openForBtn = null;
    function closePopover() {
      popover.classList.remove("is-open");
      openForBtn = null;
      document.removeEventListener("click", closeOnClickOutside);
    }
    function closeOnClickOutside(e) {
      if (popover.contains(e.target) || (e.target && e.target.closest && e.target.closest(".help-btn"))) return;
      closePopover();
    }
    app.querySelectorAll(".help-btn").forEach(function (btn) {
      btn.addEventListener("click", function (e) {
        e.stopPropagation();
        var key = btn.getAttribute("data-help");
        var text = getHelpText(key);
        if (!text) return;
        if (openForBtn === btn) {
          closePopover();
          return;
        }
        openForBtn = btn;
        popover.textContent = text;
        popover.classList.add("is-open");
        var rect = btn.getBoundingClientRect();
        var popoverWidth = 280;
        var gap = 8;
        var left = rect.right + gap;
        if (left + popoverWidth > window.innerWidth - 12) {
          left = Math.max(12, rect.left - popoverWidth - gap);
        }
        popover.style.left = left + "px";
        popover.style.top = rect.top + "px";
        setTimeout(function () {
          document.addEventListener("click", closeOnClickOutside);
        }, 0);
      });
    });

    bindSlotAndSaveButtons();

    bindCalcInputs();
    bindSensitivityInputs();
    bindOptimizerInputs();
    var printBtn = document.getElementById("finalReportPrintBtn");
    if (printBtn) printBtn.addEventListener("click", openReportPrintWindow);
    var excelBtn = document.getElementById("finalReportExcelBtn");
    if (excelBtn) excelBtn.addEventListener("click", downloadFinalReportExcel);

    refreshAll();
  }

  function bindCalcInputs() {
    var geoKwEl = document.getElementById("calcGeoKw");
    var geoCostKwEl = document.getElementById("calcGeoCostKw");
    var solarAreaEl = document.getElementById("calcSolarAreaM2");
    var roofAreaEl = document.getElementById("calcRoofAreaM2");
    var pvCostKwpEl = document.getElementById("calcPvCostKwp");
    function updateAllCalcs() {
      updateCalcGeo();
      updateCalcGeoCost();
      updateCalcSolar();
      updateCalcRoof();
      updateCalcPvCost();
    }
    if (geoKwEl) {
      geoKwEl.addEventListener("input", updateAllCalcs);
      geoKwEl.addEventListener("change", updateAllCalcs);
    }
    if (geoCostKwEl) {
      geoCostKwEl.addEventListener("input", updateAllCalcs);
      geoCostKwEl.addEventListener("change", updateAllCalcs);
    }
    if (solarAreaEl) {
      solarAreaEl.addEventListener("input", updateAllCalcs);
      solarAreaEl.addEventListener("change", updateAllCalcs);
    }
    if (roofAreaEl) {
      roofAreaEl.addEventListener("input", updateAllCalcs);
      roofAreaEl.addEventListener("change", updateAllCalcs);
    }
    if (pvCostKwpEl) {
      pvCostKwpEl.addEventListener("input", updateAllCalcs);
      pvCostKwpEl.addEventListener("change", updateAllCalcs);
    }
    updateAllCalcs();
  }

  function bindSensitivityInputs() {
    var rangeEl = document.getElementById("sensitivityRange");
    var varEl = document.getElementById("sensitivityVariable");
    function update() {
      updateSensitivityTable();
    }
    if (rangeEl) {
      rangeEl.addEventListener("input", update);
      rangeEl.addEventListener("change", update);
    }
    if (varEl) {
      varEl.addEventListener("change", update);
    }
    updateSensitivityTable();
  }

  function bindOptimizerInputs() {
    var targetEl = document.getElementById("optimizerTargetSelfSufficiency");
    if (targetEl) {
      targetEl.addEventListener("input", function () {
        inputs.optimizerTargetSelfSufficiency = targetEl.value;
      });
      targetEl.addEventListener("change", function () {
        inputs.optimizerTargetSelfSufficiency = targetEl.value;
      });
    }
    app.querySelectorAll(".optimizer-toggle-btn").forEach(function (btn) {
      btn.addEventListener("click", function () {
        var mode = btn.getAttribute("data-optimizer-mode");
        if (mode !== "investment" && mode !== "economics") return;
        inputs.optimizerMode = mode;
        app.querySelectorAll(".optimizer-toggle-btn").forEach(function (b) { b.classList.remove("is-active"); });
        btn.classList.add("is-active");
      });
    });
    var geoOptEl = document.getElementById("optimizerGeothermalOptimize");
    if (geoOptEl) {
      geoOptEl.addEventListener("change", function () {
        inputs.geothermalOptimize = geoOptEl.checked;
        var geoOptTab2 = document.getElementById("geothermalOptimize");
        if (geoOptTab2) geoOptTab2.checked = geoOptEl.checked;
      });
    }
    var runBtn = document.getElementById("optimizerRunBtn");
    if (runBtn) runBtn.addEventListener("click", runOptimizer);
  }

  function updateCalcGeo() {
    var usrtEl = document.getElementById("calcGeoUsrt");
    if (!usrtEl) return;
    var kw = parseFloat(document.getElementById("calcGeoKw").value, 10);
    if (!Number.isFinite(kw) || kw < 0) {
      usrtEl.textContent = "—";
      return;
    }
    var usrt = kw / KW_PER_USRT;
    usrtEl.textContent = usrt.toFixed(2) + " USRT";
  }

  function updateCalcGeoCost() {
    var costEl = document.getElementById("calcGeoCost");
    if (!costEl) return;
    var kw = parseFloat(document.getElementById("calcGeoCostKw").value, 10);
    if (!Number.isFinite(kw) || kw < 0) {
      costEl.textContent = "—";
      return;
    }
    costEl.textContent = formatWon(kw * GEOTHERMAL_WON_PER_KW);
  }

  function updateCalcSolar() {
    var el = document.getElementById("calcSolarKwp");
    if (!el) return;
    var area = parseFloat(document.getElementById("calcSolarAreaM2").value, 10);
    if (!Number.isFinite(area) || area < 0) {
      el.textContent = "—";
      return;
    }
    var kwp = area * 0.1;
    el.textContent = kwp.toFixed(2);
  }

  function updateCalcRoof() {
    var kwEl = document.getElementById("calcRoofKwp");
    if (!kwEl) return;
    var area = parseFloat(document.getElementById("calcRoofAreaM2").value, 10);
    if (!Number.isFinite(area) || area < 0) {
      kwEl.textContent = "—";
      return;
    }
    kwEl.textContent = (area * 0.7 * 0.1).toFixed(2);
  }

  function updateCalcPvCost() {
    var costEl = document.getElementById("calcPvCost");
    if (!costEl) return;
    var kwp = parseFloat(document.getElementById("calcPvCostKwp").value, 10);
    if (!Number.isFinite(kwp) || kwp < 0) {
      costEl.textContent = "—";
      return;
    }
    costEl.textContent = formatWon(kwp * PV_WON_PER_KW);
  }

  var MJ_PER_KWH = 3.6;

  function updateGasWonPerKwh() {
    var el = document.getElementById("gasWonPerKwh");
    if (!el) return;
    var gasMj = parseFloat(inputs.gasWonPerMj, 10);
    if (Number.isFinite(gasMj) && gasMj >= 0) {
      var wonPerKwh = gasMj * MJ_PER_KWH;
      el.textContent = wonPerKwh.toFixed(2) + " 원/kWh";
    } else {
      el.textContent = "— 원/kWh";
    }
  }

  function getTotalGasUseKwh() {
    var area = parseFloat(inputs.areaM2, 10);
    if (!Number.isFinite(area) || area < 0) return null;
    var unitGasPrimary = UNIT_ENERGY_KWH_PER_M2_YR * GAS_SHARE;
    var unitGasFinal = unitGasPrimary / GAS_PRIMARY_FACTOR;
    return unitGasFinal * area;
  }

  function getTotalElectricityUseKwh() {
    var area = parseFloat(inputs.areaM2, 10);
    if (!Number.isFinite(area) || area < 0) return null;
    var unitElecPrimary = UNIT_ENERGY_KWH_PER_M2_YR * ELECTRICITY_SHARE;
    var unitElecFinal = unitElecPrimary / ELECTRICITY_PRIMARY_FACTOR;
    return unitElecFinal * area;
  }

  function updateGasAndElectricityUsage() {
    var gasEl = document.getElementById("totalGasUseKwh");
    var gasMjEl = document.getElementById("totalGasUseMj");
    var elecEl = document.getElementById("totalElectricityUseKwh");
    var gas = getTotalGasUseKwh();
    var elec = getTotalElectricityUseKwh();
    if (gasEl) {
      gasEl.textContent = (gas != null && Number.isFinite(gas))
        ? Math.round(gas).toLocaleString("ko-KR") + " kWh/yr"
        : "—";
    }
    if (gasMjEl) {
      if (gas != null && Number.isFinite(gas)) {
        var mj = gas * MJ_PER_KWH;
        gasMjEl.textContent = "(" + Math.round(mj).toLocaleString("ko-KR") + " MJ/yr)";
      } else {
        gasMjEl.textContent = "";
      }
    }
    if (elecEl) {
      elecEl.textContent = (elec != null && Number.isFinite(elec))
        ? Math.round(elec).toLocaleString("ko-KR") + " kWh/yr"
        : "—";
    }
  }

  function updateEnergyCosts() {
    var elecSaveEl = document.getElementById("electricityCostSave");
    var gasSaveEl = document.getElementById("gasCostSave");
    var totalSaveEl = document.getElementById("totalEnergyCostSave");
    var elecKwh = getTotalElectricityUseKwh();
    var gasKwh = getTotalGasUseKwh();
    var smp = parseFloat(inputs.smpWonPerKwh, 10);
    var gasWonPerMj = parseFloat(inputs.gasWonPerMj, 10);
    var gasWonPerKwh = Number.isFinite(gasWonPerMj) ? gasWonPerMj * MJ_PER_KWH : NaN;

    var elecSave = null;
    var gasSave = null;

    /* COP 기준 세이브: 난방·급탕 가스 → 지열 히트펌프(COP 4) 대체, 냉방 전력 → 지열 히트펌프(COP 4) 대체 */
    if (gasKwh != null && Number.isFinite(gasKwh) && gasKwh >= 0 && Number.isFinite(gasWonPerKwh) && gasWonPerKwh >= 0 && Number.isFinite(smp) && smp >= 0) {
      var heatFromGas = gasKwh * BOILER_EFF;
      var heatPumpElecForHeat = heatFromGas / HEAT_PUMP_COP_HEAT;
      var gasCostCurrent = gasKwh * gasWonPerKwh;
      var elecCostForHeat = heatPumpElecForHeat * smp;
      gasSave = gasCostCurrent - elecCostForHeat;
    }
    if (elecKwh != null && Number.isFinite(elecKwh) && elecKwh >= 0 && Number.isFinite(smp) && smp >= 0) {
      var elecCool = elecKwh * COOLING_SHARE_OF_ELECTRICITY;
      var coolingLoad = elecCool * AC_COP;
      var geoCoolElec = coolingLoad / HEAT_PUMP_COP_COOL;
      var coolCostCurrent = elecCool * smp;
      var coolCostWithGeo = geoCoolElec * smp;
      elecSave = coolCostCurrent - coolCostWithGeo;
    }

    /* 절감은 지열세대에만 해당: 지열세대/총세대 비율 적용 */
    var ratio = getGeothermalHouseholdRatio();
    if (ratio != null && Number.isFinite(ratio)) {
      if (gasSave != null && Number.isFinite(gasSave)) gasSave = gasSave * ratio;
      if (elecSave != null && Number.isFinite(elecSave)) elecSave = elecSave * ratio;
    } else {
      gasSave = null;
      elecSave = null;
    }

    /* PV·BAPV·BIPV 1차에너지생산량 ÷ 2.75를 전력요금 절감에 가산 */
    var pvPrimary = calcPvPrimaryEnergyProduction();
    if (pvPrimary != null && Number.isFinite(pvPrimary) && pvPrimary >= 0 && Number.isFinite(smp) && smp >= 0) {
      var pvSaveWon = (pvPrimary / ELECTRICITY_PRIMARY_FACTOR) * smp;
      elecSave = (elecSave != null && Number.isFinite(elecSave) ? elecSave : 0) + pvSaveWon;
    }

    if (gasSaveEl) {
      gasSaveEl.textContent = (gasSave != null && Number.isFinite(gasSave)) ? formatWon(gasSave) + "/yr" : "—";
    }
    if (elecSaveEl) {
      elecSaveEl.textContent = (elecSave != null && Number.isFinite(elecSave)) ? formatWon(elecSave) + "/yr" : "—";
    }
    if (totalSaveEl) {
      var totalSave = (elecSave != null && Number.isFinite(elecSave) ? elecSave : 0) + (gasSave != null && Number.isFinite(gasSave) ? gasSave : 0);
      if (totalSave > 0 || (elecSave != null && Number.isFinite(elecSave)) || (gasSave != null && Number.isFinite(gasSave))) {
        totalSaveEl.textContent = formatWon(totalSave) + "/yr";
        updatePaybackPeriod(totalSave);
      } else {
        totalSaveEl.textContent = "—";
        updatePaybackPeriod(null);
      }
    } else {
      updatePaybackPeriod(getTotalEnergyCostSaveAmount());
    }
    updateNpvIrrReport();
  }

  function getTotalEnergyCostSaveAmount() {
    var elecKwh = getTotalElectricityUseKwh();
    var gasKwh = getTotalGasUseKwh();
    var smp = parseFloat(inputs.smpWonPerKwh, 10);
    var gasWonPerMj = parseFloat(inputs.gasWonPerMj, 10);
    var gasWonPerKwh = Number.isFinite(gasWonPerMj) ? gasWonPerMj * MJ_PER_KWH : NaN;
    var elecSave = null;
    var gasSave = null;
    if (gasKwh != null && Number.isFinite(gasKwh) && gasKwh >= 0 && Number.isFinite(gasWonPerKwh) && gasWonPerKwh >= 0 && Number.isFinite(smp) && smp >= 0) {
      var heatFromGas = gasKwh * BOILER_EFF;
      var heatPumpElecForHeat = heatFromGas / HEAT_PUMP_COP_HEAT;
      var gasCostCurrent = gasKwh * gasWonPerKwh;
      var elecCostForHeat = heatPumpElecForHeat * smp;
      gasSave = gasCostCurrent - elecCostForHeat;
    }
    if (elecKwh != null && Number.isFinite(elecKwh) && elecKwh >= 0 && Number.isFinite(smp) && smp >= 0) {
      var elecCool = elecKwh * COOLING_SHARE_OF_ELECTRICITY;
      var coolingLoad = elecCool * AC_COP;
      var geoCoolElec = coolingLoad / HEAT_PUMP_COP_COOL;
      var coolCostCurrent = elecCool * smp;
      var coolCostWithGeo = geoCoolElec * smp;
      elecSave = coolCostCurrent - coolCostWithGeo;
    }
    var ratio = getGeothermalHouseholdRatio();
    if (ratio != null && Number.isFinite(ratio)) {
      if (gasSave != null && Number.isFinite(gasSave)) gasSave = gasSave * ratio;
      if (elecSave != null && Number.isFinite(elecSave)) elecSave = elecSave * ratio;
    } else {
      gasSave = null;
      elecSave = null;
    }
    /* PV·BAPV·BIPV 1차에너지생산량 ÷ 2.75를 전력요금 절감에 가산 */
    var pvPrimary = calcPvPrimaryEnergyProduction();
    if (pvPrimary != null && Number.isFinite(pvPrimary) && pvPrimary >= 0 && Number.isFinite(smp) && smp >= 0) {
      var pvSaveWon = (pvPrimary / ELECTRICITY_PRIMARY_FACTOR) * smp;
      elecSave = (elecSave != null && Number.isFinite(elecSave) ? elecSave : 0) + pvSaveWon;
    }
    if (elecSave != null && Number.isFinite(elecSave) && gasSave != null && Number.isFinite(gasSave)) {
      return elecSave + gasSave;
    }
    /* PV만 있거나 한쪽만 있는 경우 */
    var total = (elecSave != null && Number.isFinite(elecSave) ? elecSave : 0) + (gasSave != null && Number.isFinite(gasSave) ? gasSave : 0);
    return total;
  }

  /** 전력/가스 절감 분리. overrides: { smp, gasWonPerMj } (미지정 시 inputs 사용). 민감도 분석용 */
  function getEnergySaves(overrides) {
    var o = overrides || {};
    var smp = o.smp != null && Number.isFinite(o.smp) ? o.smp : parseFloat(inputs.smpWonPerKwh, 10);
    var gasWonPerMj = o.gasWonPerMj != null && Number.isFinite(o.gasWonPerMj) ? o.gasWonPerMj : parseFloat(inputs.gasWonPerMj, 10);
    var gasWonPerKwh = Number.isFinite(gasWonPerMj) ? gasWonPerMj * MJ_PER_KWH : NaN;
    var elecKwh = getTotalElectricityUseKwh();
    var gasKwh = getTotalGasUseKwh();
    var elecSave = null;
    var gasSave = null;
    if (gasKwh != null && Number.isFinite(gasKwh) && gasKwh >= 0 && Number.isFinite(gasWonPerKwh) && gasWonPerKwh >= 0 && Number.isFinite(smp) && smp >= 0) {
      var heatFromGas = gasKwh * BOILER_EFF;
      var heatPumpElecForHeat = heatFromGas / HEAT_PUMP_COP_HEAT;
      var gasCostCurrent = gasKwh * gasWonPerKwh;
      var elecCostForHeat = heatPumpElecForHeat * smp;
      gasSave = gasCostCurrent - elecCostForHeat;
    }
    if (elecKwh != null && Number.isFinite(elecKwh) && elecKwh >= 0 && Number.isFinite(smp) && smp >= 0) {
      var elecCool = elecKwh * COOLING_SHARE_OF_ELECTRICITY;
      var coolingLoad = elecCool * AC_COP;
      var geoCoolElec = coolingLoad / HEAT_PUMP_COP_COOL;
      var coolCostCurrent = elecCool * smp;
      var coolCostWithGeo = geoCoolElec * smp;
      elecSave = coolCostCurrent - coolCostWithGeo;
    }
    var ratio = getGeothermalHouseholdRatio();
    if (ratio != null && Number.isFinite(ratio)) {
      if (gasSave != null && Number.isFinite(gasSave)) gasSave = gasSave * ratio;
      if (elecSave != null && Number.isFinite(elecSave)) elecSave = elecSave * ratio;
    } else {
      gasSave = null;
      elecSave = null;
    }
    var pvPrimary = calcPvPrimaryEnergyProduction();
    if (pvPrimary != null && Number.isFinite(pvPrimary) && pvPrimary >= 0 && Number.isFinite(smp) && smp >= 0) {
      var pvSaveWon = (pvPrimary / ELECTRICITY_PRIMARY_FACTOR) * smp;
      elecSave = (elecSave != null && Number.isFinite(elecSave) ? elecSave : 0) + pvSaveWon;
    }
    var es = elecSave != null && Number.isFinite(elecSave) ? elecSave : 0;
    var gs = gasSave != null && Number.isFinite(gasSave) ? gasSave : 0;
    return { electricitySave: es, gasSave: gs, totalSave: es + gs };
  }

  function updatePaybackPeriod(annualSave) {
    var el = document.getElementById("paybackPeriod");
    if (!el) return;
    var save = annualSave != null ? annualSave : getTotalEnergyCostSaveAmount();
    var investment = getTotalEnergyFacilityCost(inputs.geothermalOptimize, true);
    var taxSave = getAcquisitionTaxSaveAmount();
    if (taxSave == null) taxSave = 0;
    var netInvestment = (investment != null && Number.isFinite(investment)) ? investment - taxSave : null;
    if (netInvestment != null && netInvestment >= 0 && save != null && Number.isFinite(save) && save > 0) {
      var years = netInvestment / save;
      el.textContent = (years % 1 === 0 ? Math.round(years) : years.toFixed(1)) + "년";
    } else {
      el.textContent = "—";
    }
  }

  var UNIT_ENERGY_KWH_PER_M2_YR = 120;
  var GAS_SHARE = 0.625;
  var ELECTRICITY_SHARE = 0.375;
  var GAS_PRIMARY_FACTOR = 1.1;
  var ELECTRICITY_PRIMARY_FACTOR = 2.75;
  var BOILER_EFF = 0.9;
  var HEAT_PUMP_COP_HEAT = 4;
  var AC_COP = 2.5;
  var HEAT_PUMP_COP_COOL = 4;
  var COOLING_SHARE_OF_ELECTRICITY = 0.5;
  var PV_WON_PER_KW = 2068 * 1000;
  var BAPV_WON_PER_KW = PV_WON_PER_KW * 2;
  var BIPV_WON_PER_KW = PV_WON_PER_KW * 3;
  var GEOTHERMAL_WON_PER_KW = 1394 * 1000;
  var KW_PER_USRT = 3.517;
  var PIPE_WON_PER_M = {
    "100A": 3829293,
    "150A": 4455534,
    "200A": 4473246,
    "250A": 4456991,
    "300A": 4799793,
    "350A": 5054100,
    "400A": 5320712
  };

  function formatWon(n) {
    if (!Number.isFinite(n)) return "—";
    if (n < 0) return "-" + Math.round(-n).toLocaleString("ko-KR") + "원";
    return Math.round(n).toLocaleString("ko-KR") + "원";
  }

  function getHelpText(key) {
    if (!key) return "";
    if (HELP_MEMOS[key]) return HELP_MEMOS[key];
    if (key === "geothermalCapacityKw") return "지열 kW당 단가: " + formatWon(GEOTHERMAL_WON_PER_KW);
    if (key === "solarPvKwp") return "PV kWp당 단가: " + formatWon(PV_WON_PER_KW);
    if (key === "solarBapvKwp") return "BAPV kWp당 단가: " + formatWon(BAPV_WON_PER_KW);
    if (key === "solarBipvKwp") return "BIPV kWp당 단가: " + formatWon(BIPV_WON_PER_KW);
    var pipeMatch = key && key.match(/^pipeLength(100A|150A|200A|250A|300A|350A|400A)$/);
    if (pipeMatch && PIPE_WON_PER_M[pipeMatch[1]]) {
      return pipeMatch[1] + " 열수송관 m당 단가: " + formatWon(PIPE_WON_PER_M[pipeMatch[1]]);
    }
    return "";
  }

  function updateAllCosts() {
    var g = parseFloat(inputs.geothermalCapacityKw, 10) || 0;
    var pv = parseFloat(inputs.solarPvKwp, 10) || 0;
    var bapv = parseFloat(inputs.solarBapvKwp, 10) || 0;
    var bipv = parseFloat(inputs.solarBipvKwp, 10) || 0;
    var geoCost = g * GEOTHERMAL_WON_PER_KW;
    var usrt = g / KW_PER_USRT;
    if (Number.isFinite(g) && g > 0 && Number.isFinite(usrt)) {
      setCost("costGeothermal", geoCost, formatWon(geoCost) + " (" + usrt.toFixed(1) + " USRT)");
    } else {
      setCost("costGeothermal", null, "—");
    }
    setCost("costPv", pv * PV_WON_PER_KW);
    setCost("costBapv", bapv * BAPV_WON_PER_KW);
    setCost("costBipv", bipv * BIPV_WON_PER_KW);
    setCost("costPipe100A", (parseFloat(inputs.pipeLength100A, 10) || 0) * PIPE_WON_PER_M["100A"]);
    setCost("costPipe150A", (parseFloat(inputs.pipeLength150A, 10) || 0) * PIPE_WON_PER_M["150A"]);
    setCost("costPipe200A", (parseFloat(inputs.pipeLength200A, 10) || 0) * PIPE_WON_PER_M["200A"]);
    setCost("costPipe250A", (parseFloat(inputs.pipeLength250A, 10) || 0) * PIPE_WON_PER_M["250A"]);
    setCost("costPipe300A", (parseFloat(inputs.pipeLength300A, 10) || 0) * PIPE_WON_PER_M["300A"]);
    setCost("costPipe350A", (parseFloat(inputs.pipeLength350A, 10) || 0) * PIPE_WON_PER_M["350A"]);
    setCost("costPipe400A", (parseFloat(inputs.pipeLength400A, 10) || 0) * PIPE_WON_PER_M["400A"]);
    updateTotalEnergyFacilityCost();
    updateCostAfterAcquisitionTaxSave();
  }

  function getTotalEnergyFacilityCost(geothermalHalf, forEconomics) {
    var g = parseFloat(inputs.geothermalCapacityKw, 10) || 0;
    var pv = parseFloat(inputs.solarPvKwp, 10) || 0;
    var bapv = parseFloat(inputs.solarBapvKwp, 10) || 0;
    var bipv = parseFloat(inputs.solarBipvKwp, 10) || 0;
    var geoCost = g * GEOTHERMAL_WON_PER_KW;
    if (forEconomics) geoCost = geoCost * 0.75; /* 지열 시 보일러/에어컨 미설치 효과: 대체 설비 가격 1/4 가정 */
    if (geothermalHalf) geoCost = geoCost * 0.5;
    var pvCost = pv * PV_WON_PER_KW + bapv * BAPV_WON_PER_KW + bipv * BIPV_WON_PER_KW;
    var pipeCost = 0;
    ["100A", "150A", "200A", "250A", "300A", "350A", "400A"].forEach(function (key) {
      var len = parseFloat(inputs["pipeLength" + key], 10) || 0;
      pipeCost += len * PIPE_WON_PER_M[key];
    });
    return geoCost + pvCost + pipeCost;
  }

  function updateTotalEnergyFacilityCost() {
    var totalFull = getTotalEnergyFacilityCost(false, false);
    var totalTab2 = getTotalEnergyFacilityCost(inputs.geothermalOptimize, true);
    var textFull = (Number.isFinite(totalFull) && totalFull >= 0) ? formatWon(totalFull) : "—";
    var textTab2 = (Number.isFinite(totalTab2) && totalTab2 >= 0) ? formatWon(totalTab2) : "—";
    var el = document.getElementById("totalEnergyFacilityCost");
    if (el) el.textContent = textFull;
    var el2 = document.getElementById("totalEnergyFacilityCostTab2");
    if (el2) el2.textContent = textTab2;
    updatePaybackPeriod();
    updateNpvIrrReport();
  }

  function getAcquisitionTaxSaveAmount() {
    var cost = getConstructionCost();
    var rate = getSelfSufficiencyRate();
    var grade = getGradeFromSelfSufficiency(rate);
    if (cost == null || !Number.isFinite(cost) || grade == null) return null;
    var discount = TAX_DISCOUNT_BY_GRADE[grade];
    return cost * TAX_RATE_BASE * discount;
  }

  function updateCostAfterAcquisitionTaxSave() {
    var el = document.getElementById("costAfterAcquisitionTaxSave");
    if (!el) return;
    var total = getTotalEnergyFacilityCost(false, false);
    var save = getAcquisitionTaxSaveAmount();
    if (save == null) save = 0;
    if (Number.isFinite(total) && total >= 0 && Number.isFinite(save)) {
      var after = Math.max(0, total - save);
      el.textContent = formatWon(after);
    } else {
      el.textContent = "—";
    }
  }

  function setCost(id, value, displayText) {
    var el = document.getElementById(id);
    if (!el) return;
    if (displayText !== undefined) {
      el.textContent = displayText;
      return;
    }
    el.textContent = formatWon(value);
  }

  function calcPrimaryEnergyProduction() {
    var pv = parseFloat(inputs.solarPvKwp, 10) || 0;
    var bapv = parseFloat(inputs.solarBapvKwp, 10) || 0;
    var bipv = parseFloat(inputs.solarBipvKwp, 10) || 0;
    var geo = parseFloat(inputs.geothermalCapacityKw, 10) || 0;
    if (!Number.isFinite(pv)) pv = 0;
    if (!Number.isFinite(bapv)) bapv = 0;
    if (!Number.isFinite(bipv)) bipv = 0;
    if (!Number.isFinite(geo)) geo = 0;
    var kwhPerYr = geo * 638 + pv * 2000 + bapv * 1111 + bipv * 1111;
    if (kwhPerYr === 0) return null;
    return kwhPerYr;
  }

  /** PV·BAPV·BIPV만의 1차에너지생산량(kWh/yr). 전력요금 절감 가산용 */
  function calcPvPrimaryEnergyProduction() {
    var pv = parseFloat(inputs.solarPvKwp, 10) || 0;
    var bapv = parseFloat(inputs.solarBapvKwp, 10) || 0;
    var bipv = parseFloat(inputs.solarBipvKwp, 10) || 0;
    if (!Number.isFinite(pv)) pv = 0;
    if (!Number.isFinite(bapv)) bapv = 0;
    if (!Number.isFinite(bipv)) bipv = 0;
    return pv * 2000 + bapv * 1111 + bipv * 1111;
  }

  function updatePrimaryEnergyProduction() {
    var el = document.getElementById("primaryEnergyProduction");
    if (!el) return;
    var v = calcPrimaryEnergyProduction();
    if (v !== null && Number.isFinite(v)) {
      el.textContent = Math.round(v).toLocaleString("ko-KR") + " kWh/yr";
    } else {
      el.textContent = "—";
    }
  }

  function getPrimaryEnergyRequired() {
    var area = parseFloat(inputs.areaM2, 10);
    if (!Number.isFinite(area) || area < 0) return null;
    return area * UNIT_ENERGY_KWH_PER_M2_YR;
  }

  function updatePrimaryEnergyRequired() {
    var el = document.getElementById("primaryEnergyRequired");
    if (!el) return;
    var v = getPrimaryEnergyRequired();
    if (v !== null && Number.isFinite(v)) {
      el.textContent = Math.round(v).toLocaleString("ko-KR") + " kWh/yr";
    } else {
      el.textContent = "—";
    }
  }

  var USRT_PER_HOUSEHOLD_REF = 3;
  var REF_AREA_M2_PER_HOUSEHOLD = 85;

  function updateGeothermalHouseholds() {
    var geoEl = document.getElementById("geothermalHouseholds");
    var nonEl = document.getElementById("nonGeothermalHouseholds");
    if (!geoEl || !nonEl) return;
    var geoKw = parseFloat(inputs.geothermalCapacityKw, 10) || 0;
    var avgArea = parseFloat(inputs.avgDedicatedAreaM2, 10);
    var totalHouseholds = parseFloat(inputs.householdCount, 10) || 0;
    var total = Number.isFinite(totalHouseholds) && totalHouseholds >= 0 ? Math.floor(totalHouseholds) : 0;

    var geoHouseholds = 0;
    if (Number.isFinite(geoKw) && geoKw > 0 && Number.isFinite(avgArea) && avgArea > 0) {
      var geoUSRT = geoKw / KW_PER_USRT;
      geoHouseholds = Math.floor(geoUSRT * REF_AREA_M2_PER_HOUSEHOLD / (USRT_PER_HOUSEHOLD_REF * avgArea));
      if (!Number.isFinite(geoHouseholds) || geoHouseholds < 0) geoHouseholds = 0;
    }
    var nonGeo = Math.max(0, total - geoHouseholds);

    geoEl.textContent = "지열세대: " + (Number.isFinite(geoKw) && geoKw > 0 && Number.isFinite(avgArea) && avgArea > 0 ? geoHouseholds : "—");
    nonEl.textContent = "비 지열세대: " + (total > 0 || geoHouseholds > 0 ? nonGeo : "—");
  }

  /** 지열세대/총세대 비율. 절감요금에 곱할 때 사용. 총세대 0이면 null. */
  function getGeothermalHouseholdRatio() {
    var geoKw = parseFloat(inputs.geothermalCapacityKw, 10) || 0;
    var avgArea = parseFloat(inputs.avgDedicatedAreaM2, 10);
    var totalHouseholds = parseFloat(inputs.householdCount, 10) || 0;
    var total = Number.isFinite(totalHouseholds) && totalHouseholds >= 0 ? Math.floor(totalHouseholds) : 0;
    if (total <= 0) return null;
    var geoHouseholds = 0;
    if (Number.isFinite(geoKw) && geoKw > 0 && Number.isFinite(avgArea) && avgArea > 0) {
      var geoUSRT = geoKw / KW_PER_USRT;
      geoHouseholds = Math.floor(geoUSRT * REF_AREA_M2_PER_HOUSEHOLD / (USRT_PER_HOUSEHOLD_REF * avgArea));
      if (!Number.isFinite(geoHouseholds) || geoHouseholds < 0) geoHouseholds = 0;
    }
    return geoHouseholds / total;
  }

  var M2_PER_PYEONG = 3.3058;

  function updateAreaPyeong() {
    var el = document.getElementById("areaPyeong");
    if (!el) return;
    var v = parseFloat(inputs.areaM2, 10);
    if (Number.isFinite(v) && v >= 0) {
      el.textContent = (v / M2_PER_PYEONG).toFixed(1) + " 평";
    } else {
      el.textContent = "— 평";
    }
  }

  function getConstructionCost() {
    var area = parseFloat(inputs.areaM2, 10);
    if (!Number.isFinite(area) || area < 0) return null;
    return area * COST_PER_M2;
  }

  function updateConstructionCost() {
    var el = document.getElementById("constructionCost");
    if (!el) return;
    var v = getConstructionCost();
    if (v !== null && Number.isFinite(v)) {
      el.textContent = "공사비: " + formatWon(v);
    } else {
      el.textContent = "공사비: —";
    }
  }

  function getSelfSufficiencyRate() {
    var prod = calcPrimaryEnergyProduction();
    var req = getPrimaryEnergyRequired();
    if (prod == null || req == null || !Number.isFinite(prod) || !Number.isFinite(req) || req <= 0) return null;
    return (prod / req) * 100;
  }

  function getGradeFromSelfSufficiency(rate) {
    if (rate == null || !Number.isFinite(rate)) return null;
    if (rate >= 100) return 1;
    if (rate >= 80) return 2;
    if (rate >= 60) return 3;
    if (rate >= 40) return 4;
    if (rate >= 20) return 5;
    return null;
  }

  function updateSelfSufficiencyRate() {
    var el = document.getElementById("selfSufficiencyRate");
    if (!el) return;
    var rate = getSelfSufficiencyRate();
    if (rate !== null && Number.isFinite(rate)) {
      el.textContent = rate.toFixed(1) + "%";
    } else {
      el.textContent = "—";
    }
    updateReportGrade();
  }

  function updateReportGrade() {
    var el = document.getElementById("reportGradeLabel");
    if (!el) return;
    var rate = getSelfSufficiencyRate();
    var grade = getGradeFromSelfSufficiency(rate);
    if (grade !== null) {
      el.textContent = grade + "등급";
    } else {
      el.textContent = "—";
    }
  }

  var TAX_RATE_BASE = 0.0316;
  var TAX_DISCOUNT_BY_GRADE = { 1: 0.20, 2: 0.20, 3: 0.20, 4: 0.18, 5: 0.15 };

  /* NPV/IRR 프로젝션 가정: 자기자본 100%, 할인율 4.5%, 물가상승 2%, 20년, 법인세율 25% */
  var DISCOUNT_RATE = 0.045;
  var INFLATION_RATE = 0.02;
  var PROJECTION_YEARS = 20;
  var CORPORATE_TAX_RATE = 0.25;

  /** 20년 프로젝션: 투자비·취득세절감·에너지절감 기준 세전/세후 이익 및 캐시플로우 */
  function getProjectionData() {
    var investment = getTotalEnergyFacilityCost(inputs.geothermalOptimize, true);
    var taxSave = getAcquisitionTaxSaveAmount();
    var annualSave = getTotalEnergyCostSaveAmount();
    if (taxSave == null) taxSave = 0;
    if (investment == null || !Number.isFinite(investment)) investment = 0;
    if (annualSave == null || !Number.isFinite(annualSave)) annualSave = 0;

    var preTax = [];
    var tax = [];
    var afterTax = [];
    var cashFlows = [];

    var cf0 = taxSave - investment;
    cashFlows.push(cf0);
    preTax.push(cf0);
    tax.push(cf0 >= 0 ? cf0 * CORPORATE_TAX_RATE : 0);
    afterTax.push(cf0 - (cf0 >= 0 ? cf0 * CORPORATE_TAX_RATE : 0));

    for (var t = 1; t <= PROJECTION_YEARS; t++) {
      var energySaveT = annualSave * Math.pow(1 + INFLATION_RATE, t - 1);
      preTax.push(energySaveT);
      var taxT = energySaveT >= 0 ? energySaveT * CORPORATE_TAX_RATE : 0;
      tax.push(taxT);
      afterTax.push(energySaveT - taxT);
      cashFlows.push(energySaveT - taxT);
    }

    return {
      investment: investment,
      taxSave: taxSave,
      annualSave: annualSave,
      preTax: preTax,
      tax: tax,
      afterTax: afterTax,
      cashFlows: cashFlows
    };
  }

  /** 민감도 분석용: 투자비·취득세절감·연간절감을 직접 넣어 프로젝션 데이터 생성 */
  function getProjectionDataFromParams(investment, taxSave, annualSave) {
    if (investment == null || !Number.isFinite(investment)) investment = 0;
    if (taxSave == null || !Number.isFinite(taxSave)) taxSave = 0;
    if (annualSave == null || !Number.isFinite(annualSave)) annualSave = 0;
    var preTax = [];
    var tax = [];
    var afterTax = [];
    var cashFlows = [];
    var cf0 = taxSave - investment;
    cashFlows.push(cf0);
    preTax.push(cf0);
    tax.push(cf0 >= 0 ? cf0 * CORPORATE_TAX_RATE : 0);
    afterTax.push(cf0 - (cf0 >= 0 ? cf0 * CORPORATE_TAX_RATE : 0));
    for (var t = 1; t <= PROJECTION_YEARS; t++) {
      var energySaveT = annualSave * Math.pow(1 + INFLATION_RATE, t - 1);
      preTax.push(energySaveT);
      var taxT = energySaveT >= 0 ? energySaveT * CORPORATE_TAX_RATE : 0;
      tax.push(taxT);
      afterTax.push(energySaveT - taxT);
      cashFlows.push(energySaveT - taxT);
    }
    return {
      investment: investment,
      taxSave: taxSave,
      annualSave: annualSave,
      preTax: preTax,
      tax: tax,
      afterTax: afterTax,
      cashFlows: cashFlows
    };
  }

  function calcNPV(cashFlows, discountRate) {
    var npv = 0;
    for (var t = 0; t < cashFlows.length; t++) {
      npv += cashFlows[t] / Math.pow(1 + discountRate, t);
    }
    return npv;
  }

  function calcIRR(cashFlows) {
    var r = 0.05;
    for (var iter = 0; iter < 100; iter++) {
      var npv = calcNPV(cashFlows, r);
      var dnpv = 0;
      for (var t = 0; t < cashFlows.length; t++) {
        dnpv -= t * cashFlows[t] / Math.pow(1 + r, t + 1);
      }
      if (Math.abs(dnpv) < 1e-10) break;
      var rNew = r - npv / dnpv;
      if (rNew < -0.5) rNew = -0.5;
      if (rNew > 2) rNew = 2;
      if (Math.abs(rNew - r) < 1e-8) break;
      r = rNew;
    }
    return r;
  }

  /** 최적화용: 주어진 지열·PV 용량으로 시나리오 지표 계산 (inputs 일시 변경 후 복원) */
  function getScenarioMetrics(geoKw, pvKwp, bapvKwp, bipvKwp) {
    if (bapvKwp == null || !Number.isFinite(bapvKwp)) bapvKwp = 0;
    if (bipvKwp == null || !Number.isFinite(bipvKwp)) bipvKwp = 0;
    var savedGeo = inputs.geothermalCapacityKw;
    var savedPv = inputs.solarPvKwp;
    var savedBapv = inputs.solarBapvKwp;
    var savedBipv = inputs.solarBipvKwp;
    inputs.geothermalCapacityKw = geoKw;
    inputs.solarPvKwp = pvKwp;
    inputs.solarBapvKwp = bapvKwp;
    inputs.solarBipvKwp = bipvKwp;
    var investment = getTotalEnergyFacilityCost(inputs.geothermalOptimize, true);
    if (investment == null || !Number.isFinite(investment)) investment = 0;
    var taxSave = getAcquisitionTaxSaveAmount();
    if (taxSave == null || !Number.isFinite(taxSave)) taxSave = 0;
    var annualSave = getTotalEnergyCostSaveAmount();
    if (annualSave == null || !Number.isFinite(annualSave)) annualSave = 0;
    var rate = getSelfSufficiencyRate();
    var proj = getProjectionDataFromParams(investment, taxSave, annualSave);
    var npv = calcNPV(proj.cashFlows, DISCOUNT_RATE);
    var irr = calcIRR(proj.cashFlows);
    inputs.geothermalCapacityKw = savedGeo;
    inputs.solarPvKwp = savedPv;
    inputs.solarBapvKwp = savedBapv;
    inputs.solarBipvKwp = savedBipv;
    return {
      investment: investment,
      taxSave: taxSave,
      annualSave: annualSave,
      selfSufficiencyRate: rate,
      npv: npv,
      irr: irr
    };
  }

  /** 옥상 최대 PV: 간이계산기와 동일 (옥상면적 × 70% × 100 W) → kWp */
  function getMaxPvKwpFromRoof(roofAreaM2) {
    if (!Number.isFinite(roofAreaM2) || roofAreaM2 <= 0) return 0;
    return roofAreaM2 * 0.7 * 0.1;
  }

  /** 부지 최대 지열: 정사각형 부지, 간격 7m, 1기당 3 USRT */
  function getMaxGeoKwFromSite(siteAreaM2) {
    if (!Number.isFinite(siteAreaM2) || siteAreaM2 <= 0) return 0;
    var side = Math.sqrt(siteAreaM2);
    var boreholesPerSide = Math.floor(side / 7) + 1;
    var totalBoreholes = boreholesPerSide * boreholesPerSide;
    var maxUSRT = totalBoreholes * 3;
    return maxUSRT * KW_PER_USRT;
  }

  /** 자립률 목표에 맞춰 지열·PV·BAPV·BIPV 채우기 (옥상/부지 제한 반영) 후 결과 표시 */
  function runOptimizer() {
    var resultEl = document.getElementById("optimizerResult");
    var targetVal = parseFloat(document.getElementById("optimizerTargetSelfSufficiency").value, 10);
    if (!Number.isFinite(targetVal) || targetVal < 0 || targetVal > 100) {
      if (resultEl) resultEl.innerHTML = "<p class=\"optimizer-result-error\">자립률 목표를 0~100 사이 숫자로 입력해 주세요.</p>";
      return;
    }
    var areaVal = parseFloat(document.getElementById("optimizerAreaM2").value, 10);
    if (!Number.isFinite(areaVal) || areaVal <= 0) {
      if (resultEl) resultEl.innerHTML = "<p class=\"optimizer-result-error\">연면적을 입력해 주세요.</p>";
      return;
    }
    var roofVal = parseFloat(document.getElementById("optimizerRoofAreaM2").value, 10);
    var siteVal = parseFloat(document.getElementById("optimizerSiteAreaM2").value, 10);
    if (!Number.isFinite(roofVal) || roofVal < 0) roofVal = 0;
    if (!Number.isFinite(siteVal) || siteVal < 0) siteVal = 0;

    var geoOptCheck = document.getElementById("optimizerGeothermalOptimize");
    if (geoOptCheck) inputs.geothermalOptimize = geoOptCheck.checked;
    var geoOptTab2 = document.getElementById("geothermalOptimize");
    if (geoOptTab2) geoOptTab2.checked = !!inputs.geothermalOptimize;

    var householdVal = document.getElementById("optimizerHouseholdCount").value;
    var householdNum = householdVal === "" ? 0 : parseFloat(householdVal, 10);
    if (!Number.isFinite(householdNum) || householdNum < 0) householdNum = 0;

    inputs.areaM2 = areaVal;
    inputs.householdCount = householdNum;
    var areaEl = document.getElementById("areaM2");
    var householdEl = document.getElementById("householdCount");
    if (areaEl) areaEl.value = areaVal;
    if (householdEl) householdEl.value = householdNum;

    var required = getPrimaryEnergyRequired();
    if (required == null || !Number.isFinite(required) || required <= 0) {
      if (resultEl) resultEl.innerHTML = "<p class=\"optimizer-result-error\">연면적을 확인해 주세요.</p>";
      return;
    }
    var targetProduction = (targetVal / 100) * required;
    var maxPv = getMaxPvKwpFromRoof(roofVal);
    var maxGeo = getMaxGeoKwFromSite(siteVal);

    var geo = 0;
    var pv = 0;
    var bapv = 0;
    var bipv = 0;
    var mode = inputs.optimizerMode;
    var remaining = targetProduction;

    if (mode === "investment") {
      pv = Math.min(maxPv, remaining / 2000);
      if (!Number.isFinite(pv) || pv < 0) pv = 0;
      remaining -= pv * 2000;
      geo = Math.min(maxGeo, remaining / 638);
      if (!Number.isFinite(geo) || geo < 0) geo = 0;
      remaining -= geo * 638;
      if (remaining > 0) {
        bapv = remaining / 1111;
        if (!Number.isFinite(bapv) || bapv < 0) bapv = 0;
        remaining -= bapv * 1111;
        if (remaining > 0) bipv = remaining / 1111;
      }
    } else {
      geo = Math.min(maxGeo, remaining / 638);
      if (!Number.isFinite(geo) || geo < 0) geo = 0;
      remaining -= geo * 638;
      pv = Math.min(maxPv, remaining / 2000);
      if (!Number.isFinite(pv) || pv < 0) pv = 0;
      remaining -= pv * 2000;
      if (remaining > 0) {
        bapv = remaining / 1111;
        if (!Number.isFinite(bapv) || bapv < 0) bapv = 0;
        remaining -= bapv * 1111;
        if (remaining > 0) bipv = remaining / 1111;
      }
    }

    var m = getScenarioMetrics(geo, pv, bapv, bipv);
    var rate = m.selfSufficiencyRate;
    if (rate == null || !Number.isFinite(rate) || rate < targetVal - 0.5) {
      if (resultEl) resultEl.innerHTML = "<p class=\"optimizer-result-error\">자립률 " + targetVal + "%를 만족하는 조합을 찾지 못했습니다. 옥상면적·부지면적을 늘리거나 목표를 낮춰 주세요.</p>";
      return;
    }

    var html = "<div class=\"optimizer-result-box\">";
    html += "<h4 class=\"optimizer-result-title\">최적 시나리오</h4>";
    html += "<ul class=\"optimizer-result-list\">";
    html += "<li><strong>연면적</strong> " + (areaVal % 1 === 0 ? areaVal : areaVal.toFixed(1)) + " m²</li>";
    html += "<li><strong>세대수</strong> " + (householdNum % 1 === 0 ? householdNum : householdNum.toFixed(0)) + " 세대</li>";
    html += "<li><strong>지열</strong> " + (geo % 1 === 0 ? geo : geo.toFixed(1)) + " kW</li>";
    html += "<li><strong>PV</strong> " + (pv % 1 === 0 ? pv : pv.toFixed(1)) + " kWp</li>";
    if (bapv > 0.01) html += "<li><strong>BAPV</strong> " + (bapv % 1 === 0 ? bapv : bapv.toFixed(1)) + " kWp</li>";
    if (bipv > 0.01) html += "<li><strong>BIPV</strong> " + (bipv % 1 === 0 ? bipv : bipv.toFixed(1)) + " kWp</li>";
    html += "<li><strong>자립률</strong> " + (rate != null && Number.isFinite(rate) ? rate.toFixed(1) : "—") + " %</li>";
    html += "<li><strong>추가 투자비(경제성 기준)</strong> " + formatWon(m.investment) + "</li>";
    html += "<li><strong>NPV</strong> " + (Number.isFinite(m.npv) ? formatWon(Math.round(m.npv)) : "—") + "</li>";
    html += "<li><strong>IRR</strong> " + (Number.isFinite(m.irr) ? (m.irr * 100).toFixed(2) + "%" : "—") + "</li>";
    html += "</ul>";
    html += "<button type=\"button\" class=\"optimizer-apply-btn\" id=\"optimizerApplyBtn\">이 결과를 투자비산출에 반영</button>";
    html += "</div>";
    if (resultEl) resultEl.innerHTML = html;

    var applyBtn = document.getElementById("optimizerApplyBtn");
    if (applyBtn) {
      applyBtn.addEventListener("click", function () {
        inputs.areaM2 = areaVal;
        inputs.householdCount = householdNum;
        inputs.geothermalCapacityKw = geo;
        inputs.solarPvKwp = pv;
        inputs.solarBapvKwp = bapv > 0.01 ? bapv : "";
        inputs.solarBipvKwp = bipv > 0.01 ? bipv : "";
        inputs.checkGeothermal = geo > 0;
        inputs.checkPv = pv > 0;
        inputs.checkBapv = bapv > 0.01;
        inputs.checkBipv = bipv > 0.01;
        var areaInput = document.getElementById("areaM2");
        var householdInput = document.getElementById("householdCount");
        var geoEl = document.getElementById("geothermalCapacityKw");
        var pvEl = document.getElementById("solarPvKwp");
        var bapvEl = document.getElementById("solarBapvKwp");
        var bipvEl = document.getElementById("solarBipvKwp");
        var checkGeo = document.getElementById("checkGeothermal");
        var checkPv = document.getElementById("checkPv");
        var checkBapv = document.getElementById("checkBapv");
        var checkBipv = document.getElementById("checkBipv");
        var rowGeo = document.getElementById("renewableRowGeothermal");
        var rowPv = document.getElementById("renewableRowPv");
        var rowBapv = document.getElementById("renewableRowBapv");
        var rowBipv = document.getElementById("renewableRowBipv");
        if (areaInput) areaInput.value = areaVal;
        if (householdInput) householdInput.value = householdNum;
        if (geoEl) geoEl.value = geo;
        if (pvEl) pvEl.value = pv;
        if (bapvEl) bapvEl.value = bapv > 0.01 ? bapv : "";
        if (bipvEl) bipvEl.value = bipv > 0.01 ? bipv : "";
        if (checkGeo) { checkGeo.checked = geo > 0; if (rowGeo) rowGeo.style.display = geo > 0 ? "block" : "none"; }
        if (checkPv) { checkPv.checked = pv > 0; if (rowPv) rowPv.style.display = pv > 0 ? "block" : "none"; }
        if (checkBapv) { checkBapv.checked = bapv > 0.01; if (rowBapv) rowBapv.style.display = bapv > 0.01 ? "block" : "none"; }
        if (checkBipv) { checkBipv.checked = bipv > 0.01; if (rowBipv) rowBipv.style.display = bipv > 0.01 ? "block" : "none"; }
        var geoOptCheck = document.getElementById("optimizerGeothermalOptimize");
        if (geoOptCheck) inputs.geothermalOptimize = geoOptCheck.checked;
        var geoOptTab2 = document.getElementById("geothermalOptimize");
        if (geoOptTab2) geoOptTab2.checked = !!inputs.geothermalOptimize;
        refreshAll();
      });
    }
  }

  function updateNpvIrrReport() {
    var data = getProjectionData();
    var cashFlows = data.cashFlows;
    var npvEl = document.getElementById("npvValue");
    var irrEl = document.getElementById("irrValue");
    var tbody = document.getElementById("projectionTableBody");
    var assumptionsEl = document.getElementById("npvIrrAssumptions");

    if (assumptionsEl) {
      assumptionsEl.innerHTML = "자기자본 100% · 할인율 " + (DISCOUNT_RATE * 100) + "% · 물가상승 " + (INFLATION_RATE * 100) + "% · " + PROJECTION_YEARS + "년 · 세율(세전이익 기준) " + (CORPORATE_TAX_RATE * 100) + "%";
    }
    if (npvEl) {
      var npv = calcNPV(cashFlows, DISCOUNT_RATE);
      npvEl.textContent = Number.isFinite(npv) ? formatWon(npv) : "—";
    }
    if (irrEl) {
      var irr = calcIRR(cashFlows);
      irrEl.textContent = Number.isFinite(irr) ? (irr * 100).toFixed(2) + "%" : "—";
    }
    if (tbody) {
      var rows = "";
      for (var y = 0; y <= PROJECTION_YEARS; y++) {
        var label = y === 0 ? "당해년도" : y + "년차";
        var pt = data.preTax[y];
        var tx = data.tax[y];
        var at = data.afterTax[y];
        var pvAt = Number.isFinite(at) ? at / Math.pow(1 + DISCOUNT_RATE, y) : NaN;
        var npvStr = Number.isFinite(pvAt) ? formatWon(Math.round(pvAt)) : "—";
        rows += "<tr><td>" + label + "</td><td>" + (Number.isFinite(pt) ? Math.round(pt).toLocaleString("ko-KR") : "—") + "</td><td>" + (Number.isFinite(tx) ? Math.round(tx).toLocaleString("ko-KR") : "—") + "</td><td>" + (Number.isFinite(at) ? Math.round(at).toLocaleString("ko-KR") : "—") + "</td><td>" + npvStr + "</td></tr>";
      }
      tbody.innerHTML = rows;
    }
  }

  function updateSensitivityTable() {
    var rangeEl = document.getElementById("sensitivityRange");
    var varEl = document.getElementById("sensitivityVariable");
    if (!rangeEl || !varEl) return;
    var range = parseFloat(rangeEl.value, 10);
    if (!Number.isFinite(range) || range < 1 || range > 50) range = 10;
    var variable = varEl.value;
    var mults = [
      1 - range / 100,
      1 - range / 200,
      1,
      1 + range / 200,
      1 + range / 100
    ];
    var pctLabels = [-range, -range / 2, 0, range / 2, range];
    for (var c = 0; c < 5; c++) {
      var colEl = document.getElementById("sensitivityCol" + c);
      if (colEl) colEl.textContent = pctLabels[c] + "%";
    }
    var inv0 = getTotalEnergyFacilityCost(inputs.geothermalOptimize, true);
    var tax0 = getAcquisitionTaxSaveAmount();
    if (tax0 == null || !Number.isFinite(tax0)) tax0 = 0;
    var smp0 = parseFloat(inputs.smpWonPerKwh, 10);
    var gas0 = parseFloat(inputs.gasWonPerMj, 10);
    var baseSaves = getEnergySaves({});
    var empty = "—";
    var npvValues = [];
    var irrValues = [];
    function setCell(id, text) {
      var el = document.getElementById(id);
      if (el) el.textContent = text;
    }
    for (var i = 0; i < 5; i++) {
      var m = mults[i];
      var inv = (inv0 != null && Number.isFinite(inv0)) ? inv0 : 0;
      var tax = tax0;
      var smp = Number.isFinite(smp0) ? smp0 : 0;
      var gas = Number.isFinite(gas0) ? gas0 : 0;
      if (variable === "investment") {
        inv = inv * m;
      } else if (variable === "smp") {
        smp = smp * m;
      } else if (variable === "gas") {
        gas = gas * m;
      } else if (variable === "taxSave") {
        tax = tax * m;
      }
      var saves = baseSaves;
      if (variable === "smp" && Number.isFinite(smp0)) {
        saves = getEnergySaves({ smp: smp });
      } else if (variable === "gas" && Number.isFinite(gas0)) {
        saves = getEnergySaves({ gasWonPerMj: gas });
      }
      var totalSave = saves.totalSave;
      var netInv = inv - tax;
      var payback = (totalSave != null && Number.isFinite(totalSave) && totalSave > 0 && netInv >= 0)
        ? (netInv / totalSave)
        : null;
      var proj = getProjectionDataFromParams(inv, tax, totalSave);
      var npv = calcNPV(proj.cashFlows, DISCOUNT_RATE);
      var irr = calcIRR(proj.cashFlows);
      npvValues.push(Number.isFinite(npv) ? npv : null);
      irrValues.push(Number.isFinite(irr) ? irr * 100 : null);
      setCell("sensElec" + i, saves.electricitySave != null && Number.isFinite(saves.electricitySave) ? formatWon(Math.round(saves.electricitySave)) : empty);
      setCell("sensGas" + i, saves.gasSave != null && Number.isFinite(saves.gasSave) ? formatWon(Math.round(saves.gasSave)) : empty);
      setCell("sensTotal" + i, totalSave != null && Number.isFinite(totalSave) ? formatWon(Math.round(totalSave)) : empty);
      setCell("sensPayback" + i, payback != null && Number.isFinite(payback) ? (payback % 1 === 0 ? Math.round(payback) : payback.toFixed(1)) + "년" : empty);
      setCell("sensNpv" + i, Number.isFinite(npv) ? formatWon(Math.round(npv)) : empty);
      setCell("sensIrr" + i, Number.isFinite(irr) ? (irr * 100).toFixed(2) + "%" : empty);
    }
    updateSensitivityCharts(pctLabels, npvValues, irrValues);
  }

  function updateSensitivityCharts(labels, npvValues, irrValues) {
    var npvWrap = document.getElementById("sensitivityNpvChart");
    var irrWrap = document.getElementById("sensitivityIrrChart");
    var npvWrapFinal = document.getElementById("finalSensitivityNpvChart");
    var irrWrapFinal = document.getElementById("finalSensitivityIrrChart");
    if (!npvWrap && !irrWrap && !npvWrapFinal && !irrWrapFinal) return;
    var n = labels.length;
    var w = 320;
    var h = 200;
    var pad = { top: 20, right: 20, bottom: 36, left: 52 };
    var chartW = w - pad.left - pad.right;
    var chartH = h - pad.top - pad.bottom;

    function scaleNpv() {
      var dataMin = Infinity;
      var dataMax = -Infinity;
      for (var i = 0; i < n; i++) {
        var v = npvValues[i];
        if (v != null && Number.isFinite(v)) {
          if (v < dataMin) dataMin = v;
          if (v > dataMax) dataMax = v;
        }
      }
      if (dataMin === Infinity) dataMin = 0;
      if (dataMax === -Infinity) dataMax = 0;
      var axisMax = dataMax * 1.2;
      var axisMin = dataMin - dataMax * 1.2;
      if (dataMax === 0 && dataMin === 0) {
        axisMin = -1;
        axisMax = 1;
      } else if (axisMin >= axisMax) {
        axisMin = dataMin - Math.abs(dataMax - dataMin) || dataMin - 1;
        axisMax = dataMax + Math.abs(dataMax - dataMin) || dataMax + 1;
      }
      return {
        min: axisMin,
        max: axisMax,
        toY: function (v) {
          if (v == null || !Number.isFinite(v)) return pad.top + chartH / 2;
          var t = (v - axisMin) / (axisMax - axisMin);
          return pad.top + chartH * (1 - t);
        }
      };
    }

    function scaleIrr() {
      var dataMin = Infinity;
      var dataMax = -Infinity;
      for (var i = 0; i < n; i++) {
        var v = irrValues[i];
        if (v != null && Number.isFinite(v)) {
          if (v < dataMin) dataMin = v;
          if (v > dataMax) dataMax = v;
        }
      }
      if (dataMin === Infinity) dataMin = 0;
      if (dataMax === -Infinity) dataMax = 10;
      var axisMax = dataMax * 1.2;
      var axisMin = dataMin - dataMax * 1.2;
      if (dataMax === 0 && dataMin === 0) {
        axisMin = 0;
        axisMax = 10;
      } else if (axisMin >= axisMax) {
        var range = Math.abs(dataMax - dataMin) || 5;
        axisMin = dataMin - range;
        axisMax = dataMax + range;
      }
      return {
        min: axisMin,
        max: axisMax,
        toY: function (v) {
          if (v == null || !Number.isFinite(v)) return pad.top + chartH / 2;
          var t = (v - axisMin) / (axisMax - axisMin);
          return pad.top + chartH * (1 - t);
        }
      };
    }

    var npvScale = scaleNpv();
    var irrScale = scaleIrr();
    var barW = Math.max(12, (chartW / n) * 0.5);
    var gap = chartW / n;
    var barX = function (i) { return pad.left + (i + 0.5) * gap - barW / 2; };

    var zeroY = (0 >= npvScale.min && 0 <= npvScale.max)
      ? npvScale.toY(0)
      : (0 >= npvScale.max ? pad.top : pad.top + chartH);

    var svgNpv = '<svg class="sensitivity-svg" viewBox="0 0 ' + w + ' ' + h + '" preserveAspectRatio="xMidYMid meet">';
    svgNpv += '<line x1="' + pad.left + '" y1="' + (pad.top + chartH) + '" x2="' + (w - pad.right) + '" y2="' + (pad.top + chartH) + '" stroke="rgba(255,255,255,0.2)" stroke-width="1"/>';
    svgNpv += '<line x1="' + pad.left + '" y1="' + pad.top + '" x2="' + pad.left + '" y2="' + (pad.top + chartH) + '" stroke="rgba(255,255,255,0.2)" stroke-width="1"/>';
    if (zeroY > pad.top && zeroY < pad.top + chartH) {
      svgNpv += '<line x1="' + pad.left + '" y1="' + zeroY + '" x2="' + (w - pad.right) + '" y2="' + zeroY + '" stroke="rgba(255,255,255,0.15)" stroke-width="1" stroke-dasharray="4,2"/>';
    }
    for (var i = 0; i < n; i++) {
      var v = npvValues[i];
      var yVal = npvScale.toY(v);
      var by = Math.min(yVal, zeroY);
      var bh = Math.abs(yVal - zeroY);
      if (v != null && Number.isFinite(v)) {
        var fill = "var(--color-1)";
        svgNpv += '<rect x="' + barX(i) + '" y="' + by + '" width="' + barW + '" height="' + bh + '" fill="' + fill + '" fill-opacity="0.85" rx="2"/>';
      }
      svgNpv += '<text x="' + (pad.left + (i + 0.5) * gap) + '" y="' + (h - 8) + '" text-anchor="middle" fill="var(--text-soft)" font-size="11">' + labels[i] + '%</text>';
    }
    var npvMin = npvScale.min;
    var npvMax = npvScale.max;
    if (Number.isFinite(npvMin) && Number.isFinite(npvMax)) {
      svgNpv += '<text x="' + (pad.left - 6) + '" y="' + (pad.top + 4) + '" text-anchor="end" fill="var(--text-soft)" font-size="10">' + (npvMax >= 1e8 ? (npvMax / 1e8).toFixed(1) + "억" : Math.round(npvMax).toLocaleString("ko-KR")) + '</text>';
      svgNpv += '<text x="' + (pad.left - 6) + '" y="' + (pad.top + chartH) + '" text-anchor="end" fill="var(--text-soft)" font-size="10">' + (npvMin >= 1e8 || npvMin <= -1e8 ? (npvMin / 1e8).toFixed(1) + "억" : Math.round(npvMin).toLocaleString("ko-KR")) + '</text>';
    }
    svgNpv += '</svg>';
    if (npvWrap) npvWrap.innerHTML = svgNpv;
    if (npvWrapFinal) npvWrapFinal.innerHTML = svgNpv;

    var svgIrr = '<svg class="sensitivity-svg" viewBox="0 0 ' + w + ' ' + h + '" preserveAspectRatio="xMidYMid meet">';
    svgIrr += '<line x1="' + pad.left + '" y1="' + (pad.top + chartH) + '" x2="' + (w - pad.right) + '" y2="' + (pad.top + chartH) + '" stroke="rgba(255,255,255,0.2)" stroke-width="1"/>';
    svgIrr += '<line x1="' + pad.left + '" y1="' + pad.top + '" x2="' + pad.left + '" y2="' + (pad.top + chartH) + '" stroke="rgba(255,255,255,0.2)" stroke-width="1"/>';
    var irrPath = [];
    for (var i = 0; i < n; i++) {
      var v = irrValues[i];
      var y = irrScale.toY(v);
      if (v != null && Number.isFinite(v)) {
        var rx = pad.left + (i + 0.5) * gap;
        irrPath.push(rx + "," + y);
      }
    }
    if (irrPath.length > 1) {
      svgIrr += '<polyline points="' + irrPath.join(" ") + '" fill="none" stroke="var(--color-1)" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>';
    }
    for (var i = 0; i < n; i++) {
      var v = irrValues[i];
      var y = irrScale.toY(v);
      if (v != null && Number.isFinite(v)) {
        svgIrr += '<circle cx="' + (pad.left + (i + 0.5) * gap) + '" cy="' + y + '" r="5" fill="var(--color-1)" stroke="rgba(255,255,255,0.4)" stroke-width="1"/>';
      }
      svgIrr += '<text x="' + (pad.left + (i + 0.5) * gap) + '" y="' + (h - 8) + '" text-anchor="middle" fill="var(--text-soft)" font-size="11">' + labels[i] + '%</text>';
    }
    var irrMin = irrScale.min;
    var irrMax = irrScale.max;
    if (Number.isFinite(irrMin) && Number.isFinite(irrMax)) {
      svgIrr += '<text x="' + (pad.left - 6) + '" y="' + (pad.top + 4) + '" text-anchor="end" fill="var(--text-soft)" font-size="10">' + irrMax.toFixed(1) + '%</text>';
      svgIrr += '<text x="' + (pad.left - 6) + '" y="' + (pad.top + chartH) + '" text-anchor="end" fill="var(--text-soft)" font-size="10">' + irrMin.toFixed(1) + '%</text>';
    }
    svgIrr += '</svg>';
    if (irrWrap) irrWrap.innerHTML = svgIrr;
    if (irrWrapFinal) irrWrapFinal.innerHTML = svgIrr;
  }

  function updateFinalReport() {
    var invBody = document.getElementById("finalReportInvestmentBody");
    var econBody = document.getElementById("finalReportEconomicsBody");
    var sensBody = document.getElementById("finalReportSensitivityBody");
    var sensNote = document.getElementById("finalReportSensitivityNote");
    if (!invBody || !econBody || !sensBody) return;

    function pair(label, value) {
      if (!label && (value == null || value === "")) {
        return "<td class=\"final-report-label\"></td><td class=\"final-report-value\"></td>";
      }
      var v = value != null && value !== "" ? value : "—";
      return "<td class=\"final-report-label\">" + (label || "") + "</td><td class=\"final-report-value\">" + v + "</td>";
    }

    function row2(label1, value1, label2, value2) {
      return "<tr>" + pair(label1, value1) + pair(label2, value2) + "</tr>";
    }

    var area = parseFloat(inputs.areaM2, 10);
    var constructionCost = getConstructionCost();
    var primaryReq = getPrimaryEnergyRequired();
    var primaryProd = calcPrimaryEnergyProduction();
    var selfRate = getSelfSufficiencyRate();
    var grade = getGradeFromSelfSufficiency(selfRate);
    var taxSave = getAcquisitionTaxSaveAmount();
    var totalInv = getTotalEnergyFacilityCost(false, false);
    var costAfterTax = totalInv != null && taxSave != null && Number.isFinite(totalInv) && Number.isFinite(taxSave) ? Math.max(0, totalInv - taxSave) : null;

    var areaStr = Number.isFinite(area) && area >= 0 ? Math.round(area).toLocaleString("ko-KR") + " m²" : null;
    var hhStr = inputs.householdCount ? inputs.householdCount + " 세대" : null;
    var avgAreaStr = inputs.avgDedicatedAreaM2 ? inputs.avgDedicatedAreaM2 + " m²" : null;
    var costStr = constructionCost != null && Number.isFinite(constructionCost) ? formatWon(constructionCost) : null;
    var geoStr = inputs.checkGeothermal && inputs.geothermalCapacityKw ? inputs.geothermalCapacityKw + " kW" : "—";
    var pvStr = inputs.checkPv && inputs.solarPvKwp ? inputs.solarPvKwp + " kWp" : "—";
    var bapvStr = inputs.checkBapv && inputs.solarBapvKwp ? inputs.solarBapvKwp + " kWp" : "—";
    var bipvStr = inputs.checkBipv && inputs.solarBipvKwp ? inputs.solarBipvKwp + " kWp" : "—";
    var primaryReqStr = primaryReq != null && Number.isFinite(primaryReq) ? Math.round(primaryReq).toLocaleString("ko-KR") + " kWh/yr" : null;
    var primaryProdStr = primaryProd != null && Number.isFinite(primaryProd) ? Math.round(primaryProd).toLocaleString("ko-KR") + " kWh/yr" : null;
    var selfRateStr = selfRate != null && Number.isFinite(selfRate) ? selfRate.toFixed(1) + "%" : null;
    var gradeStr = grade != null ? grade + "등급" : null;
    var taxSaveStr = taxSave != null && Number.isFinite(taxSave) ? formatWon(taxSave) : null;
    var totalInvStr = totalInv != null && Number.isFinite(totalInv) ? formatWon(totalInv) : null;
    var costAfterTaxStr = costAfterTax != null && Number.isFinite(costAfterTax) ? formatWon(costAfterTax) : null;

    var invHtml = "";
    invHtml += row2("연면적", areaStr, "세대수", hhStr);
    invHtml += row2("평균 전용면적", avgAreaStr, "", "");
    invHtml += row2("지열 설치용량", geoStr, "PV 설치용량", pvStr);
    invHtml += row2("BAPV 설치용량", bapvStr, "BIPV 설치용량", bipvStr);
    invHtml += row2("1차에너지 소요량", primaryReqStr, "1차에너지생산량", primaryProdStr);
    invHtml += row2("자립률", selfRateStr, "ZEB 등급", gradeStr);
    invBody.innerHTML = invHtml;

    var invTab2 = getTotalEnergyFacilityCost(inputs.geothermalOptimize, true);
    var annualSave = getTotalEnergyCostSaveAmount();
    var elecUse = getTotalElectricityUseKwh();
    var gasUse = getTotalGasUseKwh();
    var taxSaveNum = getAcquisitionTaxSaveAmount();
    if (taxSaveNum == null) taxSaveNum = 0;
    var netInv = (invTab2 != null && Number.isFinite(invTab2)) ? invTab2 - taxSaveNum : null;
    var payback = (annualSave != null && Number.isFinite(annualSave) && annualSave > 0 && netInv != null && netInv >= 0) ? (netInv / annualSave) : null;
    var proj = getProjectionData();
    var npv = proj && proj.cashFlows ? calcNPV(proj.cashFlows, DISCOUNT_RATE) : null;
    var irr = proj && proj.cashFlows ? calcIRR(proj.cashFlows) : null;
    var saves = getEnergySaves({});

    var econHtml = "";
    function cell(cls, text, allowEmpty) {
      var v = (text != null && text !== "") ? text : (allowEmpty ? "" : "—");
      return "<td class=\"" + cls + "\">" + v + "</td>";
    }
    function row4(l1, v1, l2, v2) {
      return "<tr>" + cell("final-report-label", l1, true) + cell("final-report-value", v1, false) + cell("final-report-label", l2, true) + cell("final-report-value", v2, false) + "</tr>";
    }
    econHtml += row4(
      "추가 에너지설비 투자비",
      invTab2 != null && Number.isFinite(invTab2) ? formatWon(invTab2) : null,
      "취득세 절감",
      taxSaveNum !== 0 && Number.isFinite(taxSaveNum) ? formatWon(taxSaveNum) : null
    );
    econHtml += row4(
      "연간 전력 사용량",
      elecUse != null && Number.isFinite(elecUse) ? Math.round(elecUse).toLocaleString("ko-KR") + " kWh/yr" : null,
      "연간 가스 사용량",
      gasUse != null && Number.isFinite(gasUse) ? Math.round(gasUse).toLocaleString("ko-KR") + " kWh/yr" : null
    );
    econHtml += row4(
      "전기 단가",
      inputs.smpWonPerKwh ? Number(inputs.smpWonPerKwh).toLocaleString("ko-KR") + " 원/kWh" : null,
      "가스 단가",
      inputs.gasWonPerMj ? Number(inputs.gasWonPerMj).toLocaleString("ko-KR") + " 원/MJ" : null
    );
    econHtml += row4(
      "전력요금 절감",
      saves.electricitySave != null && Number.isFinite(saves.electricitySave) ? formatWon(Math.round(saves.electricitySave)) + "/yr" : null,
      "가스요금 절감",
      saves.gasSave != null && Number.isFinite(saves.gasSave) ? formatWon(Math.round(saves.gasSave)) + "/yr" : null
    );
    econHtml += "<tr>" + cell("final-report-label", "총 에너지비용 절감", true) + cell("final-report-value", annualSave != null && Number.isFinite(annualSave) ? formatWon(Math.round(annualSave)) + "/yr" : null, false) + cell("final-report-label", "간이 투자회수기간", true) + cell("final-report-value", payback != null && Number.isFinite(payback) ? (payback % 1 === 0 ? Math.round(payback) : payback.toFixed(1)) + "년" : null, false) + "</tr>";
    econHtml += "<tr>" + cell("final-report-label", "NPV", true) + cell("final-report-value", npv != null && Number.isFinite(npv) ? formatWon(Math.round(npv)) : null, false) + cell("final-report-label", "IRR", true) + cell("final-report-value", irr != null && Number.isFinite(irr) ? (irr * 100).toFixed(2) + "%" : null, false) + "</tr>";
    econBody.innerHTML = econHtml;

    var rangeEl = document.getElementById("sensitivityRange");
    var varEl = document.getElementById("sensitivityVariable");
    var range = (rangeEl && Number.isFinite(parseFloat(rangeEl.value, 10))) ? parseFloat(rangeEl.value, 10) : 10;
    var variable = varEl ? varEl.value : "investment";
    var varNames = { investment: "투자비", smp: "전기세", gas: "가스요금", taxSave: "취득세절감" };
    var pctLabels = [-range, -range / 2, 0, range / 2, range];
    for (var c = 0; c < 5; c++) {
      var colEl = document.getElementById("finalSensCol" + c);
      if (colEl) colEl.textContent = pctLabels[c] + "%";
    }
    if (sensNote) sensNote.textContent = "변수: " + (varNames[variable] || variable) + ", 범위: ±" + range + "%";

    var mults = [1 - range / 100, 1 - range / 200, 1, 1 + range / 200, 1 + range / 100];
    var inv0 = getTotalEnergyFacilityCost(inputs.geothermalOptimize, true);
    var tax0 = getAcquisitionTaxSaveAmount();
    if (tax0 == null) tax0 = 0;
    var smp0 = parseFloat(inputs.smpWonPerKwh, 10);
    var gas0 = parseFloat(inputs.gasWonPerMj, 10);
    var baseSaves = getEnergySaves({});
    var sensRows = [];
    var sensRowNames = ["전력요금 절감", "가스요금 절감", "총 에너지비용 절감", "간이투자회수 기간", "NPV", "IRR"];
    for (var i = 0; i < 5; i++) {
      var m = mults[i];
      var inv = (inv0 != null && Number.isFinite(inv0)) ? inv0 : 0;
      var tax = tax0;
      var smp = Number.isFinite(smp0) ? smp0 : 0;
      var gas = Number.isFinite(gas0) ? gas0 : 0;
      if (variable === "investment") inv = inv * m;
      else if (variable === "smp") smp = smp * m;
      else if (variable === "gas") gas = gas * m;
      else if (variable === "taxSave") tax = tax * m;
      var savesI = baseSaves;
      if (variable === "smp" && Number.isFinite(smp0)) savesI = getEnergySaves({ smp: smp });
      else if (variable === "gas" && Number.isFinite(gas0)) savesI = getEnergySaves({ gasWonPerMj: gas });
      var totalSave = savesI.totalSave;
      var netInvI = inv - tax;
      var paybackI = (totalSave != null && Number.isFinite(totalSave) && totalSave > 0 && netInvI >= 0) ? (netInvI / totalSave) : null;
      var projI = getProjectionDataFromParams(inv, tax, totalSave);
      var npvI = calcNPV(projI.cashFlows, DISCOUNT_RATE);
      var irrI = calcIRR(projI.cashFlows);
      sensRows.push({
        elec: savesI.electricitySave != null && Number.isFinite(savesI.electricitySave) ? formatWon(Math.round(savesI.electricitySave)) : "—",
        gas: savesI.gasSave != null && Number.isFinite(savesI.gasSave) ? formatWon(Math.round(savesI.gasSave)) : "—",
        total: totalSave != null && Number.isFinite(totalSave) ? formatWon(Math.round(totalSave)) : "—",
        payback: paybackI != null && Number.isFinite(paybackI) ? (paybackI % 1 === 0 ? Math.round(paybackI) : paybackI.toFixed(1)) + "년" : "—",
        npv: Number.isFinite(npvI) ? formatWon(Math.round(npvI)) : "—",
        irr: Number.isFinite(irrI) ? (irrI * 100).toFixed(2) + "%" : "—"
      });
    }
    var sensHtml = "";
    for (var r = 0; r < 6; r++) {
      sensHtml += "<tr><td class=\"final-report-label\">" + sensRowNames[r] + "</td>";
      for (var i = 0; i < 5; i++) {
        var key = ["elec", "gas", "total", "payback", "npv", "irr"][r];
        sensHtml += "<td class=\"final-report-value\">" + sensRows[i][key] + "</td>";
      }
      sensHtml += "</tr>";
    }
    sensBody.innerHTML = sensHtml;
  }

  function openReportPrintWindow() {
    updateFinalReport();
    var reportEl = document.querySelector(".final-report");
    if (!reportEl) return;
    var printContent = "";
    for (var i = 0; i < reportEl.children.length; i++) {
      var c = reportEl.children[i];
      if (c.classList && c.classList.contains("final-report-actions")) continue;
      printContent += c.outerHTML;
    }
    var styleStr = ":root{--bg-deep:#fff;--color-1:#0066cc;--text:#111;--text-soft:#333;}";
    styleStr += "body{background:#fff;color:#111;font-family:'Pretendard',sans-serif;margin:24px;line-height:1.6;}";
    styleStr += ".final-report-main-title{font-size:1.5rem;font-weight:800;margin:0 0 8px 0;color:#111;}";
    styleStr += ".final-report-desc{color:#555;font-size:0.9rem;margin:0 0 24px 0;}";
    styleStr += ".final-report-section{margin-bottom:32px;}";
    styleStr += ".final-report-section-title{font-size:1.05rem;font-weight:700;color:#0066cc;margin:0 0 14px 0;padding-bottom:8px;border-bottom:1px solid #ccc;}";
    styleStr += ".final-report-section-head{display:flex;align-items:center;justify-content:space-between;margin-bottom:14px;}";
    styleStr += ".final-report-section-head .final-report-section-title{margin:0;padding-bottom:8px;border-bottom:1px solid #ccc;}";
    styleStr += ".final-report-copy-btn{display:none;}";
    styleStr += ".final-report-table-wrap{background:#f9f9f9;border-radius:8px;border:1px solid #ddd;padding:4px;overflow-x:auto;}";
    styleStr += ".final-report-table{width:100%;border-collapse:collapse;font-size:0.9rem;}";
    styleStr += ".final-report-table th,.final-report-table td{padding:10px 14px;text-align:left;border-bottom:1px solid #e0e0e0;}";
    styleStr += ".final-report-table thead th{font-weight:700;color:#111;background:#eee;border-bottom:1px solid #ccc;}";
    styleStr += ".final-report-label{font-weight:600;color:#333;white-space:nowrap;}";
    styleStr += ".final-report-value{color:#111;font-weight:500;}";
    styleStr += ".final-report-note{font-size:0.8rem;color:#555;margin:10px 0 0 0;}";
    styleStr += ".sensitivity-charts{display:grid;grid-template-columns:1fr 1fr;gap:24px;margin-top:20px;}";
    styleStr += ".sensitivity-chart-box{background:#f9f9f9;border-radius:8px;border:1px solid #ddd;padding:16px;}";
    styleStr += ".sensitivity-chart-title{font-size:0.95rem;font-weight:700;color:#111;margin:0 0 12px 0;}";
    styleStr += ".sensitivity-chart-svg-wrap{width:100%;min-height:200px;}";
    styleStr += ".sensitivity-svg{width:100%;height:auto;min-height:200px;}";
    var fullHtml = "<!DOCTYPE html><html lang=\"ko\"><head><meta charset=\"utf-8\"><title>최종보고서</title><style>" + styleStr + "</style></head><body><div class=\"final-report\">" + printContent + "</div></body></html>";
    var win = window.open("", "_blank", "width=900,height=900,scrollbars=yes,resizable=yes");
    if (!win) return;
    win.document.open();
    win.document.write(fullHtml);
    win.document.close();
  }

  function downloadFinalReportExcel() {
    if (typeof XLSX === "undefined") {
      alert("엑셀 내보내기를 사용할 수 없습니다. (XLSX 라이브러리 로드 필요)");
      return;
    }
    try {
      updateFinalReport();

      function tableBodyToRows(tbody) {
        if (!tbody) return [];
        var rows = [];
        var trs = tbody.querySelectorAll("tr");
        for (var r = 0; r < trs.length; r++) {
          var tds = trs[r].querySelectorAll("td");
          var row = [];
          for (var c = 0; c < tds.length; c++) row.push((tds[c].textContent || "").trim());
          rows.push(row);
        }
        return rows;
      }

      var invBody = document.getElementById("finalReportInvestmentBody");
      var econBody = document.getElementById("finalReportEconomicsBody");
      var sensBody = document.getElementById("finalReportSensitivityBody");
      var sensTable = sensBody ? sensBody.closest("table") : null;
      var sensHeader = [];
      if (sensTable && sensTable.tHead) {
        var ths = sensTable.tHead.querySelectorAll("th");
        for (var i = 0; i < ths.length; i++) sensHeader.push((ths[i].textContent || "").trim());
      }
      var invRows = tableBodyToRows(invBody);
      var econRows = tableBodyToRows(econBody);
      var sensRows = tableBodyToRows(sensBody);

      // --- 총괄 화면 시트 (셀 병합 포함) ---
      var ov = [];
      var merges = [];
      var r = 0;

      function addPairRow(l1, v1, l2, v2) {
        ov.push([l1 || "", v1 || "", "", l2 || "", v2 || "", ""]);
        merges.push({s:{r:r,c:1}, e:{r:r,c:2}});
        if (l2 || v2) merges.push({s:{r:r,c:4}, e:{r:r,c:5}});
        r++;
      }
      function addEmptyRow() { ov.push(["","","","","",""]); r++; }

      for (var i = 0; i < invRows.length; i++) {
        var row = invRows[i];
        if (i === 1) {
          addPairRow(row[0], row[1], "에너지믹스 방향", "");
        } else {
          addPairRow(row[0], row[1], row[2], row[3]);
        }
        if (i === 1) addEmptyRow();
      }
      addEmptyRow();
      for (var i = 0; i < econRows.length; i++) {
        var row = econRows[i];
        addPairRow(row[0], row[1], row[2], row[3]);
      }
      addEmptyRow();
      if (sensHeader.length) { ov.push(sensHeader); r++; }
      for (var i = 0; i < sensRows.length; i++) { ov.push(sensRows[i]); r++; }

      var rangeEl = document.getElementById("sensitivityRange");
      var varEl = document.getElementById("sensitivityVariable");
      var sensRange = (rangeEl && Number.isFinite(parseFloat(rangeEl.value, 10))) ? parseFloat(rangeEl.value, 10) : 10;
      var sensVar = varEl ? varEl.value : "investment";
      var sensVarNames = { investment: "투자비", smp: "전기세", gas: "가스요금", taxSave: "취득세절감" };
      addEmptyRow();
      addPairRow("민감도 분석 변수", sensVarNames[sensVar] || sensVar, "범위", "±" + sensRange + "%");

      var wsOv = XLSX.utils.aoa_to_sheet(ov);
      wsOv["!merges"] = merges;
      wsOv["!cols"] = [
        {wch:22},{wch:22},{wch:16},{wch:20},{wch:22},{wch:16}
      ];

      var wsInv = XLSX.utils.aoa_to_sheet([["항목", "값", "항목", "값"]].concat(invRows));
      var wsEcon = XLSX.utils.aoa_to_sheet([["항목", "값", "항목", "값"]].concat(econRows));
      var sensData = sensHeader.length ? [sensHeader] : [];
      sensData = sensData.concat(sensRows);
      var wsSens = XLSX.utils.aoa_to_sheet(sensData);

      var wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, wsOv, "총괄");
      XLSX.utils.book_append_sheet(wb, wsInv, "투자비 산출 요약");
      XLSX.utils.book_append_sheet(wb, wsEcon, "경제성 분석 요약");
      XLSX.utils.book_append_sheet(wb, wsSens, "민감도 분석 요약");

      var now = new Date();
      var yy = String(now.getFullYear()).slice(-2);
      var mm = String(now.getMonth() + 1).padStart(2, "0");
      var dd = String(now.getDate()).padStart(2, "0");
      var hh = String(now.getHours()).padStart(2, "0");
      var mi = String(now.getMinutes()).padStart(2, "0");
      var ss = String(now.getSeconds()).padStart(2, "0");
      var filename = yy + mm + dd + "_최종보고서(" + hh + mi + ss + ").xlsx";

      XLSX.writeFile(wb, filename);

      setTimeout(function () {
        alert("엑셀 파일이 다운로드되었습니다.\n다운로드 폴더에서 '" + filename + "' 파일을 열어주세요.");
      }, 300);
    } catch (e) {
      alert("엑셀 저장 중 오류가 발생했습니다: " + (e && e.message ? e.message : String(e)));
    }
  }

  function copyFinalReportSectionToClipboard(btn) {
    if (typeof html2canvas === "undefined") {
      alert("이미지 복사 기능을 사용할 수 없습니다. (html2canvas 로드 필요)");
      return;
    }
    var section = btn && btn.closest && btn.closest(".final-report-section");
    if (!section) return;
    var sectionNum = parseInt(btn.getAttribute("data-section"), 10);
    if (!Number.isFinite(sectionNum) || sectionNum < 1 || sectionNum > 3) return;

    var mmToPx = function (mm) { return (mm / 25.4) * 96; };
    var targetW = Math.round(mmToPx(270));
    var targetH1 = Math.round(mmToPx(80));

    btn.disabled = true;
    html2canvas(section, {
      useCORS: true,
      allowTaint: true,
      scale: 2,
      backgroundColor: null,
      ignoreElements: function (el) {
        return el.classList && el.classList.contains("final-report-copy-btn");
      }
    }).then(function (canvas) {
      var w = canvas.width;
      var h = canvas.height;
      var outW = targetW;
      var outH = sectionNum === 1 ? targetH1 : Math.max(1, Math.round((h / w) * targetW));
      var out = document.createElement("canvas");
      out.width = outW;
      out.height = outH;
      var ctx = out.getContext("2d");
      ctx.fillStyle = "#0a0e14";
      ctx.fillRect(0, 0, outW, outH);
      ctx.drawImage(canvas, 0, 0, w, h, 0, 0, outW, outH);
      return out.toBlob("image/png");
    }).then(function (blob) {
      if (!blob) return Promise.reject(new Error("Blob 생성 실패"));
      return navigator.clipboard.write([new ClipboardItem({ "image/png": blob })]);
    }).then(function () {
      alert("클립보드에 이미지가 복사되었습니다.");
    }).catch(function (err) {
      alert("복사에 실패했습니다: " + (err && err.message ? err.message : "알 수 없는 오류"));
    }).finally(function () {
      btn.disabled = false;
    });
  }

  document.addEventListener("click", function (e) {
    var btn = e.target && e.target.closest && e.target.closest(".final-report-copy-btn");
    if (!btn) return;
    e.preventDefault();
    copyFinalReportSectionToClipboard(btn);
  });

  function updateAcquisitionTaxSave() {
    var cost = getConstructionCost();
    var rate = getSelfSufficiencyRate();
    var grade = getGradeFromSelfSufficiency(rate);
    var text = "—";
    if (cost != null && Number.isFinite(cost) && grade != null) {
      var discount = TAX_DISCOUNT_BY_GRADE[grade];
      var save = cost * TAX_RATE_BASE * discount;
      text = formatWon(save);
    }
    var el = document.getElementById("acquisitionTaxSave");
    if (el) el.textContent = text;
    var el2 = document.getElementById("acquisitionTaxSaveTab2");
    if (el2) el2.textContent = text;
    updatePaybackPeriod();
  }

  function updateGrade() {
    // 등급은 자립률 기반으로 취득세 계산에만 사용 (화면 배지 제거됨)
  }

  render();
})();
