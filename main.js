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
    [["saveBtnTab1", "saveSlotTab1"], ["saveBtnTab2", "saveSlotTab2"]].forEach(function (pair) {
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
          <h1>Seoul Energy Platform</h1>
          <div class="tab-buttons">
            <button type="button" class="tab-btn ${currentTab === 1 ? "is-active" : ""}" data-tab="1">투자비산출</button>
            <button type="button" class="tab-btn ${currentTab === 2 ? "is-active" : ""}" data-tab="2">경제성분석</button>
          </div>
          <div class="header-slots-wrap">
            <div class="developer-slot-buttons" aria-label="개발자 프리셋">
              ${[1,2,3,4,5,6,7,8,9,10].map(function(n) { return "<button type=\"button\" class=\"dev-slot-btn\" data-dev-slot=\"" + n + "\" aria-label=\"개발자 슬롯 " + n + "\">" + n + "</button>"; }).join("")}
            </div>
            <div class="slot-buttons" aria-label="저장 슬롯">
              ${[1,2,3,4,5,6,7,8,9,10].map(function(n) { return "<button type=\"button\" class=\"slot-btn\" data-slot=\"" + n + "\" aria-label=\"슬롯 " + n + " 불러오기\">" + n + "</button>"; }).join("")}
            </div>
          </div>
        </div>
        <div id="tabPanel1" class="tab-panel ${currentTab === 1 ? "is-active" : ""}">
          <p class="sub">태양광·지열 투자에 따른 자립률, 절감비용, 취득세 경감 계산</p>

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
        <div class="save-row">
          <label class="save-row-label">저장할 슬롯</label>
          <select id="saveSlotTab1" class="save-slot-select" aria-label="저장할 슬롯">
            ${[1,2,3,4,5,6,7,8,9,10].map(function(n) { return "<option value=\"" + n + "\">" + n + "</option>"; }).join("")}
          </select>
          <button type="button" id="saveBtnTab1" class="save-submit-btn">저장</button>
        </div>
        </div>

        <div id="tabPanel2" class="tab-panel ${currentTab === 2 ? "is-active" : ""}">
          <p class="sub">경제성 분석</p>
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
        <div class="save-row">
          <label class="save-row-label">저장할 슬롯</label>
          <select id="saveSlotTab2" class="save-slot-select" aria-label="저장할 슬롯">
            ${[1,2,3,4,5,6,7,8,9,10].map(function(n) { return "<option value=\"" + n + "\">" + n + "</option>"; }).join("")}
          </select>
          <button type="button" id="saveBtnTab2" class="save-submit-btn">저장</button>
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

    refreshAll();
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
