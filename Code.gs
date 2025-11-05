// Google Apps Script - Code.gs
// 분말 검사 시스템 서버사이드 로직

function doGet(e) {
  try {
    // 디버깅: 요청된 페이지 확인
    const page = e.parameter.page || 'index';
    
    // 페이지별 처리 전 디버깅 HTML 반환
    if (page === 'inspection') {
      // inspection 페이지 요청 시 디버깅 정보 표시
      return HtmlService.createHtmlOutput(`
        <html>
        <body>
          <h1>Inspection 페이지 디버깅</h1>
          <p>요청된 페이지: ${page}</p>
          <p>현재 시간: ${new Date().toLocaleString()}</p>
          <p>파라미터: ${JSON.stringify(e.parameter)}</p>
          <hr>
          <h2>파일 존재 여부 확인:</h2>
          <div id="fileCheck">확인 중...</div>
          <br>
          <a href="?">대시보드로 돌아가기</a>
          
          <script>
            // sessionStorage 데이터 확인
            const inspectionData = sessionStorage.getItem('inspectionData');
            const fileCheckDiv = document.getElementById('fileCheck');
            
            if (inspectionData) {
              try {
                const data = JSON.parse(inspectionData);
                fileCheckDiv.innerHTML = 
                  '<p style="color: green;">✅ SessionStorage 데이터 존재</p>' +
                  '<p>분말명: ' + data.data.powderName + '</p>' +
                  '<p>검사 항목 수: ' + data.items.length + '개</p>';
              } catch (error) {
                fileCheckDiv.innerHTML = '<p style="color: red;">❌ SessionStorage 데이터 파싱 실패</p>';
              }
            } else {
              fileCheckDiv.innerHTML = '<p style="color: orange;">⚠️ SessionStorage 데이터 없음</p>';
            }
          </script>
        </body>
        </html>
      `).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    
    // 다른 페이지들은 기존 로직
    let template;
    switch(page) {
      case 'search':
        template = HtmlService.createTemplateFromFile('search');
        break;
      default:
        template = HtmlService.createTemplateFromFile('index');
        break;
    }
    
    return template.evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      
  } catch (error) {
    // 오류 발생 시 디버깅 정보 표시
    return HtmlService.createHtmlOutput(`
      <html>
      <body>
        <h1>doGet 함수 오류 발생</h1>
        <p><strong>요청된 페이지:</strong> ${e.parameter.page || 'index'}</p>
        <p><strong>오류 내용:</strong> ${error.toString()}</p>
        <p><strong>오류 스택:</strong> ${error.stack}</p>
        <p><a href="?">대시보드로 돌아가기</a></p>
      </body>
      </html>
    `).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// 시트 초기화 및 설정
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 분말사양 시트 생성
  let specSheet = ss.getSheetByName('PowderSpec');
  if (!specSheet) {
    specSheet = ss.insertSheet('PowderSpec');
    const headers = [
      'PowderName', 'FlowRateMin', 'FlowRateMax', 'FlowRateType',
      'ApparentDensityMin', 'ApparentDensityMax', 'ApparentDensityType',
      'CContentMin', 'CContentMax', 'CContentType',
      'CuContentMin', 'CuContentMax', 'CuContentType',
      'NiContentMin', 'NiContentMax', 'NiContentType',
      'MoContentMin', 'MoContentMax', 'MoContentType',
      'SinterChangeRateMin', 'SinterChangeRateMax', 'SinterChangeRateType',
      'SinterStrengthMin', 'SinterStrengthMax', 'SinterStrengthType',
      'FormingStrengthMin', 'FormingStrengthMax', 'FormingStrengthType',
      'FormingLoadMin', 'FormingLoadMax', 'FormingLoadType',
      'ParticleSizeType'
    ];
    specSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // 샘플 데이터 추가 (Type 컬럼 포함)
    const sampleData = [
      ['ABC-100', 25, 35, '일상', 2.5, 3.0, '일상', 0.5, 0.8, '일상', 1.0, 2.0, '정기', '', '', '비활성', 0.3, 0.5, '일상', 5, 8, '정기', 850, 950, '정기', 120, 150, '정기', 180, 220, '정기', '정기'],
      ['DEF-200', 20, 30, '정기', 2.0, 2.8, '일상', 0.4, 0.7, '일상', '', '', '비활성', 3.0, 5.0, '정기', '', '', '비활성', 4, 7, '일상', 800, 900, '정기', 100, 140, '정기', 160, 200, '정기', '정기'],
      ['GHI-300', 30, 40, '일상', 3.0, 3.5, '일상', 0.6, 0.9, '일상', 2.0, 3.0, '일상', 1.0, 2.0, '일상', 0.2, 0.4, '정기', 6, 9, '정기', 900, 1000, '정기', 130, 160, '정기', 200, 240, '정기', '정기']
    ];
    specSheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
  }
  
  // 검사결과 시트 생성
  // 검사결과 시트 생성
  let resultSheet = ss.getSheetByName('InspectionResult');
  if (!resultSheet) {
    resultSheet = ss.insertSheet('InspectionResult');
    const headers = [
      'PowderName', 'LotNumber', 'Inspector', 'InspectionTime', 'InspectionType',
      'FlowRate1', 'FlowRate2', 'FlowRate3', 'FlowRateAvg', 'FlowRateResult',
      'ApparentDensity_EmptyCup1', 'ApparentDensity_PowderWeight1', 'ApparentDensity1',
      'ApparentDensity_EmptyCup2', 'ApparentDensity_PowderWeight2', 'ApparentDensity2',
      'ApparentDensity_EmptyCup3', 'ApparentDensity_PowderWeight3', 'ApparentDensity3',
      'ApparentDensityAvg', 'ApparentDensityResult',
      'CContent1', 'CContent2', 'CContent3', 'CContentAvg', 'CContentResult',
      'CuContent1', 'CuContent2', 'CuContent3', 'CuContentAvg', 'CuContentResult',
      'NiContent1', 'NiContent2', 'NiContent3', 'NiContentAvg', 'NiContentResult',
      'MoContent1', 'MoContent2', 'MoContent3', 'MoContentAvg', 'MoContentResult',
      'SinterChangeRate1', 'SinterChangeRate2', 'SinterChangeRate3', 'SinterChangeRateAvg', 'SinterChangeRateResult',
      'SinterStrength1', 'SinterStrength2', 'SinterStrength3', 'SinterStrengthAvg', 'SinterStrengthResult',
      'FormingStrength1', 'FormingStrength2', 'FormingStrength3', 'FormingStrengthAvg', 'FormingStrengthResult',
      'FormingLoad1', 'FormingLoad2', 'FormingLoad3', 'FormingLoadAvg', 'FormingLoadResult',
      'ParticleSize80_1', 'ParticleSize80_2', 'ParticleSize80_Avg', 'ParticleSize80_Result',
      'ParticleSize100_1', 'ParticleSize100_2', 'ParticleSize100_Avg', 'ParticleSize100_Result',
      'ParticleSize150_1', 'ParticleSize150_2', 'ParticleSize150_Avg', 'ParticleSize150_Result',
      'ParticleSize200_1', 'ParticleSize200_2', 'ParticleSize200_Avg', 'ParticleSize200_Result',
      'ParticleSize325_1', 'ParticleSize325_2', 'ParticleSize325_Avg', 'ParticleSize325_Result',
      'ParticleSize325M_1', 'ParticleSize325M_2', 'ParticleSize325M_Avg', 'ParticleSize325M_Result',
      'ParticleSizeResult',
      'FinalResult'
    ];
    resultSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  
  // 진행중검사 시트 생성
  let progressSheet = ss.getSheetByName('InspectionProgress');
  if (!progressSheet) {
    progressSheet = ss.insertSheet('InspectionProgress');
    const headers = ['PowderName', 'LotNumber', 'InspectionType', 'Inspector', 'StartTime', 'CompletedItems', 'TotalItems', 'Progress'];
    progressSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    progressSheet.getRange('H:H').setNumberFormat('@');
  }
  
  // 입도분석 규격 시트 생성
  let particleSizeSheet = ss.getSheetByName('ParticleSize');
  if (!particleSizeSheet) {
    particleSizeSheet = ss.insertSheet('ParticleSize');
    const headers = ['PowderName', 'MeshSize', 'MinValue', 'MaxValue'];
    particleSizeSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // 샘플 데이터 추가
    const sampleParticleData = [
      ['ABC-100', '+80#', 5.0, 10.0],
      ['ABC-100', '+100#', 10.0, 15.0],
      ['ABC-100', '+150#', 15.0, 20.0],
      ['ABC-100', '+200#', 20.0, 25.0],
      ['ABC-100', '+325#', 15.0, 20.0],
      ['ABC-100', '-325#', 10.0, 15.0],
      ['DEF-200', '+80#', 3.0, 8.0],
      ['DEF-200', '+100#', 8.0, 12.0],
      ['DEF-200', '+150#', 12.0, 18.0],
      ['DEF-200', '+200#', 18.0, 23.0],
      ['DEF-200', '+325#', 12.0, 18.0],
      ['DEF-200', '-325#', 8.0, 12.0]
    ];
    particleSizeSheet.getRange(2, 1, sampleParticleData.length, sampleParticleData[0].length).setValues(sampleParticleData);
  }

  // Inspector 시트 생성
  let inspectorSheet = ss.getSheetByName('Inspector');
  if (!inspectorSheet) {
    inspectorSheet = ss.insertSheet('Inspector');
    const headers = ['InspectorName'];
    inspectorSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // 샘플 데이터 추가
    const sampleInspectors = [
      ['김철수'],
      ['이영희'],
      ['박민수'],
      ['정수진']
    ];
    inspectorSheet.getRange(2, 1, sampleInspectors.length, sampleInspectors[0].length).setValues(sampleInspectors);
  }
  
  
  Logger.log('시트 초기화 완료');
  return '시트 초기화 완료';
}

// 분말 사양 조회
function getPowderSpec(powderName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('PowderSpec');
    
    if (!sheet) {
      Logger.log('PowderSpec 시트를 찾을 수 없습니다.');
      return null;
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === powderName) {
        const spec = {};
        for (let j = 0; j < headers.length; j++) {
          spec[headers[j]] = data[i][j];
        }
        return spec;
      }
    }
    return null;
  } catch (error) {
    Logger.log('getPowderSpec 오류: ' + error.toString());
    return null;
  }
}

// Inspector 목록 조회 함수 추가
function getInspectorList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Inspector');
    
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const inspectors = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        inspectors.push(data[i][0]);
      }
    }
    
    return inspectors;
    
  } catch (error) {
    Logger.log('getInspectorList 오류: ' + error.toString());
    return [];
  }
}

// 검사 항목 필터링 (PowderSpec Type 컬럼 방식)
function getInspectionItems(powderName, inspectionType) {
  try {
    const spec = getPowderSpec(powderName);
    if (!spec) return [];
    
    const allItems = [
      {name: 'FlowRate', displayName: '유동도', columns: ['FlowRateMin', 'FlowRateMax'], typeColumn: 'FlowRateType', unit: 's/50g'},
      {name: 'ApparentDensity', displayName: '겉보기밀도', columns: ['ApparentDensityMin', 'ApparentDensityMax'], typeColumn: 'ApparentDensityType', unit: 'g/cm³'},
      {name: 'CContent', displayName: 'C함량', columns: ['CContentMin', 'CContentMax'], typeColumn: 'CContentType', unit: '%'},
      {name: 'CuContent', displayName: 'Cu함량', columns: ['CuContentMin', 'CuContentMax'], typeColumn: 'CuContentType', unit: '%'},
      {name: 'NiContent', displayName: 'Ni함량', columns: ['NiContentMin', 'NiContentMax'], typeColumn: 'NiContentType', unit: '%'},
      {name: 'MoContent', displayName: 'Mo함량', columns: ['MoContentMin', 'MoContentMax'], typeColumn: 'MoContentType', unit: '%'},
      {name: 'SinterChangeRate', displayName: '소결변화율', columns: ['SinterChangeRateMin', 'SinterChangeRateMax'], typeColumn: 'SinterChangeRateType', unit: '%'},
      {name: 'SinterStrength', displayName: '소결강도', columns: ['SinterStrengthMin', 'SinterStrengthMax'], typeColumn: 'SinterStrengthType', unit: 'MPa'},
      {name: 'FormingStrength', displayName: '성형강도', columns: ['FormingStrengthMin', 'FormingStrengthMax'], typeColumn: 'FormingStrengthType', unit: 'N'},
      {name: 'FormingLoad', displayName: '성형하중', columns: ['FormingLoadMin', 'FormingLoadMax'], typeColumn: 'FormingLoadType', unit: 'kN'},
      {name: 'ParticleSize', displayName: '입도분석', columns: [], typeColumn: 'ParticleSizeType', unit: '%', isParticleSize: true}
    ];
    
    const filteredItems = allItems.filter(item => {
      const itemType = spec[item.typeColumn];
      
      // 검사 타입에 따른 필터링
      if (inspectionType === '일상점검') {
        return itemType === '일상';
      } else if (inspectionType === '정기점검') {
        return itemType === '일상' || itemType === '정기';
      }
      
      return false;
    }).filter(item => {
      // 입도분석은 별도 처리
      if (item.isParticleSize) {
        const particleSpecs = getParticleSizeSpec(powderName);
        if (particleSpecs.length > 0) {
          item.particleSpecs = particleSpecs;
          return true;
        }
        return false;
      }
      
      // 기존 규격 확인 로직
      const hasSpec = item.columns.some(col => {
        const value = spec[col];
        return value !== undefined && value !== null && value !== '' && value !== '-';
      });
      
      if (hasSpec) {
        item.min = spec[item.columns[0]] || '';
        item.max = spec[item.columns[1]] || '';
      }
      
      return hasSpec;
    });
    
    return filteredItems;
  } catch (error) {
    Logger.log('getInspectionItems 오류: ' + error.toString());
    return [];
  }
}

// 미완료 검사 목록 조회
function getIncompleteInspections() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('InspectionProgress');
    
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return []; // 헤더만 있는 경우
    
    const inspections = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) { // 분말명이 있는 경우
        try {
          const completedItems = data[i][5] ? JSON.parse(data[i][5]) : [];
          const totalItems = data[i][6] ? JSON.parse(data[i][6]) : [];
          const startTime = data[i][4]
            ? Utilities.formatDate(new Date(data[i][4]), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss")
            : "";

          inspections.push({
            powderName: data[i][0],
            lotNumber: data[i][1],
            inspectionType: data[i][2],
            inspector: data[i][3],
            startTime: startTime,
            completedItems: completedItems,
            totalItems: totalItems,
            progress: data[i][7]
          });
        } catch (parseError) {
          Logger.log(`[getIncompleteInspections] 파싱 오류 발생. 행: ${i+1}, 'completedItems' 값: '${data[i][5]}', 오류: ${parseError.toString()}`);
          continue;
        }
      }
    }
    
    return inspections;
  } catch (error) {
    Logger.log('getIncompleteInspections 오류: ' + error.toString());
    return [];
  }
}

// 기존 검사 조회
function getExistingInspection(powderName, lotNumber) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('InspectionProgress');
    
    if (!sheet) return null;
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === powderName && String(data[i][1]) === String(lotNumber)) {
        try {
          const startTime = data[i][4]
            ? Utilities.formatDate(new Date(data[i][4]), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss")
            : "";

          return {
            powderName: data[i][0],
            lotNumber: data[i][1],
            inspectionType: data[i][2],
            inspector: data[i][3],
            startTime: startTime,
            completedItems: data[i][5] ? JSON.parse(data[i][5]) : [],
            totalItems: data[i][6] ? JSON.parse(data[i][6]) : [],
            progress: String(data[i][7] || '0/0')  // 문자열로 강제 변환
          };
        } catch (parseError) {
          Logger.log('기존 검사 JSON 파싱 오류: ' + parseError.toString());
          return null;
        }
      }
    }
    
    return null;
  } catch (error) {
    Logger.log('getExistingInspection 오류: ' + error.toString());
    return null;
  }
}

// 새 검사 시작 또는 기존 검사 이어서 하기
function startInspection(powderName, lotNumber, inspectionType, inspector) {
  try {
    lotNumber = String(lotNumber);
    Logger.log(`=== startInspection 시작 ===`);
    Logger.log(`입력 파라미터: powderName=${powderName}, lotNumber=${lotNumber}, inspectionType=${inspectionType}, inspector=${inspector}`);
    
    // 입력 검증
    if (!powderName || !lotNumber || !inspectionType || !inspector) {
      Logger.log('입력 파라미터 누락');
      return {
        success: false,
        message: '필수 입력 항목이 누락되었습니다.'
      };
    }
    
    // 진행중인 검사가 있는지 확인
    Logger.log('기존 검사 확인 중...');
    const existingInspection = getExistingInspection(powderName, lotNumber);
    
    if (existingInspection) {
      Logger.log('기존 검사 발견:', JSON.stringify(existingInspection));
      const items = getInspectionItems(powderName, existingInspection.inspectionType);
      Logger.log(`기존 검사 항목 수: ${items.length}`);
      
      return {
        success: true,
        isExisting: true,
        data: existingInspection,
        items: items
      };
    }
    
    // 분말 사양 확인
    Logger.log('분말 사양 확인 중...');
    const spec = getPowderSpec(powderName);
    if (!spec) {
      Logger.log(`분말 사양을 찾을 수 없음: ${powderName}`);
      return {
        success: false,
        message: `해당 분말의 사양을 찾을 수 없습니다: ${powderName}`
      };
    }
    Logger.log('분말 사양 확인 완료');
    
    // 검사 항목 조회
    Logger.log('검사 항목 조회 중...');
    const items = getInspectionItems(powderName, inspectionType);
    Logger.log(`필터링된 검사 항목 수: ${items.length}`);
    
    if (items.length === 0) {
      Logger.log('검사 항목이 없음');
      return {
        success: false,
        message: '해당 분말의 검사 항목이 없습니다.'
      };
    }
    
    // 항목 상세 로그
    items.forEach(item => {
      Logger.log(`검사 항목: ${item.displayName} (${item.name}), 규격: ${item.min} ~ ${item.max} ${item.unit}`);
    });
    
    // 진행중검사 시트에 새 검사 추가
    Logger.log('진행중검사 시트에 새 검사 추가...');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const progressSheet = ss.getSheetByName('InspectionProgress');
    
    if (!progressSheet) {
      Logger.log('InspectionProgress 시트를 찾을 수 없음');
      return {
        success: false,
        message: 'InspectionProgress 시트를 찾을 수 없습니다.'
      };
    }
    
    const itemNames = items.map(item => item.name);
    Logger.log('저장할 항목 이름들:', JSON.stringify(itemNames));
    
    progressSheet.appendRow([
      powderName,
      lotNumber,
      inspectionType,
      inspector,
      new Date(),
      JSON.stringify([]),
      JSON.stringify(itemNames),
      `0/${itemNames.length}`
    ]);
    
    Logger.log('진행중검사 시트 추가 완료');
    
    const resultData = {
      success: true,
      isExisting: false,
      data: {
        powderName: powderName,
        lotNumber: lotNumber,
        inspectionType: inspectionType,
        inspector: inspector,
        completedItems: [],
        totalItems: itemNames
      },
      items: items
    };
    
    Logger.log('결과 데이터:', JSON.stringify(resultData));
    Logger.log(`=== startInspection 완료 ===`);
    
    return resultData;
    
  } catch (error) {
    Logger.log(`startInspection 오류: ${error.toString()}`);
    Logger.log(`오류 스택: ${error.stack}`);
    return {
      success: false,
      message: '검사 시작 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}

// 검사 항목 저장
function saveInspectionItem(powderName, lotNumber, itemName, values) {
  try {
    Logger.log(`항목 저장 요청: ${powderName}, ${lotNumber}, ${itemName}, ${values}`);

    // 겉보기밀도 전용 저장 처리 추가
    if (itemName === 'ApparentDensity') {
      return saveApparentDensityItem(powderName, lotNumber, values);
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 평균 계산
    const validValues = values.filter(v => v !== '' && !isNaN(v)).map(v => parseFloat(v));
    if (validValues.length === 0) {
      return { success: false, message: '유효한 측정값이 없습니다.' };
    }
    
    const average = validValues.reduce((sum, val) => sum + val, 0) / validValues.length;
    
    // 규격 확인 및 합격/불합격 판정
    const existingInspection = getExistingInspection(powderName, lotNumber);
    if (!existingInspection) {
      return { success: false, message: '진행중인 검사를 찾을 수 없습니다.' };
    }
    
    const items = getInspectionItems(powderName, existingInspection.inspectionType);
    const currentItem = items.find(item => item.name === itemName);
    
    let result = 'PASS';
    if (currentItem) {
      if (currentItem.min !== '' && average < parseFloat(currentItem.min)) {
        result = 'FAIL';
      }
      if (currentItem.max !== '' && average > parseFloat(currentItem.max)) {
        result = 'FAIL';
      }
    }
    
    // InspectionResult 시트에 저장
    const resultSheet = ss.getSheetByName('InspectionResult');
    const existingRowIndex = findExistingResultRow(powderName, lotNumber);
    
    if (existingRowIndex > 0) {
      // 기존 행 업데이트
      updateInspectionResult(existingRowIndex, itemName, values, average, result);
    } else {
      // 새 행 추가
      createNewInspectionResult(powderName, lotNumber, itemName, values, average, result);
    }
    
    // 진행중검사 시트 업데이트
    updateInspectionProgress(powderName, lotNumber, itemName);
    
    Logger.log('항목 저장 완료');
    
    return {
      success: true,
      average: average.toFixed(3),
      result: result
    };
    
  } catch (error) {
    Logger.log('saveInspectionItem 오류: ' + error.toString());
    return {
      success: false,
      message: '저장 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}

// 기존 검사 결과 행 찾기
function findExistingResultRow(powderName, lotNumber) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('InspectionResult');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === powderName && String(data[i][1]) === String(lotNumber)) {
        return i + 1; // 1-based row index
      }
    }
    return -1;
  } catch (error) {
    Logger.log('findExistingResultRow 오류: ' + error.toString());
    return -1;
  }
}

// 검사 결과 업데이트
function updateInspectionResult(rowIndex, itemName, values, average, result) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('InspectionResult');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // 겉보기밀도 전용 처리
    if (itemName === 'ApparentDensity') {
      // values 배열: [emptyCup1, powderWeight1, emptyCup2, powderWeight2, emptyCup3, powderWeight3, average]
      
      // 원본 데이터와 계산된 값 모두 저장
      for (let i = 0; i < 3; i++) {
        const emptyCup = values[i * 2];
        const powderWeight = values[i * 2 + 1];
        
        // 빈컵중량 저장
        const emptyCupCol = headers.indexOf('ApparentDensity_EmptyCup' + (i + 1)) + 1;
        if (emptyCupCol > 0) {
          sheet.getRange(rowIndex, emptyCupCol).setValue(emptyCup || '');
        }
        
        // 분말중량 저장
        const powderWeightCol = headers.indexOf('ApparentDensity_PowderWeight' + (i + 1)) + 1;
        if (powderWeightCol > 0) {
          sheet.getRange(rowIndex, powderWeightCol).setValue(powderWeight || '');
        }
        
        // 계산된 겉보기밀도 저장
        const densityCol = headers.indexOf('ApparentDensity' + (i + 1)) + 1;
        if (densityCol > 0) {
          if (emptyCup !== '' && powderWeight !== '') {
            const emptyCupNum = parseFloat(emptyCup);
            const powderWeightNum = parseFloat(powderWeight);
            if (!isNaN(emptyCupNum) && !isNaN(powderWeightNum)) {
              const apparentDensity = (powderWeightNum - emptyCupNum) / 25;
              sheet.getRange(rowIndex, densityCol).setValue(apparentDensity);
            } else {
              sheet.getRange(rowIndex, densityCol).setValue('');
            }
          } else {
            sheet.getRange(rowIndex, densityCol).setValue('');
          }
        }
      }
      
      // 평균 및 결과 저장
      const avgCol = headers.indexOf('ApparentDensityAvg') + 1;
      const resultCol = headers.indexOf('ApparentDensityResult') + 1;
      
      if (avgCol > 0) sheet.getRange(rowIndex, avgCol).setValue(average.toFixed(3));
      if (resultCol > 0) sheet.getRange(rowIndex, resultCol).setValue(result);
      
      return;
    }
    
    // 기존 로직 (다른 항목들)
    const val1Col = headers.indexOf(itemName + '1') + 1;
    const val2Col = headers.indexOf(itemName + '2') + 1;
    const val3Col = headers.indexOf(itemName + '3') + 1;
    const avgCol = headers.indexOf(itemName + 'Avg') + 1;
    const resultCol = headers.indexOf(itemName + 'Result') + 1;
    
    if (val1Col > 0) sheet.getRange(rowIndex, val1Col).setValue(values[0] || '');
    if (val2Col > 0) sheet.getRange(rowIndex, val2Col).setValue(values[1] || '');
    if (val3Col > 0) sheet.getRange(rowIndex, val3Col).setValue(values[2] || '');
    if (avgCol > 0) sheet.getRange(rowIndex, avgCol).setValue(average.toFixed(3));
    if (resultCol > 0) sheet.getRange(rowIndex, resultCol).setValue(result);
  } catch (error) {
    Logger.log('updateInspectionResult 오류: ' + error.toString());
  }
}

// 새 검사 결과 행 생성
function createNewInspectionResult(powderName, lotNumber, itemName, values, average, result) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('InspectionResult');
    const existingInspection = getExistingInspection(powderName, lotNumber);
    
    // 새 행 데이터 준비
    const newRow = Array(sheet.getLastColumn()).fill('');
    newRow[0] = powderName;
    newRow[1] = lotNumber;
    newRow[2] = existingInspection.inspector;
    newRow[3] = new Date();
    newRow[4] = existingInspection.inspectionType;
    
    sheet.appendRow(newRow);
    
    // 항목 데이터 업데이트
    const newRowIndex = sheet.getLastRow();
    updateInspectionResult(newRowIndex, itemName, values, average, result);
  } catch (error) {
    Logger.log('createNewInspectionResult 오류: ' + error.toString());
  }
}

// 진행중검사 시트 업데이트
function updateInspectionProgress(powderName, lotNumber, itemName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('InspectionProgress');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === powderName && String(data[i][1]) === String(lotNumber)) {
        
        let completedItems;
        try {
          completedItems = JSON.parse(data[i][5]);
        } catch (e) {
          // JSON 파싱에 실패하면 빈 배열로 초기화
          Logger.log(`JSON 파싱 오류로 인해 completedItems를 빈 배열로 초기화합니다. 기존 값: ${data[i][5]}`);
          completedItems = [];
        }

        const totalItems = data[i][6] ? JSON.parse(data[i][6]) : [];
        
        if (!completedItems.includes(itemName)) {
          completedItems.push(itemName);
        }
        
        const progress = `${completedItems.length}/${totalItems.length}`;
        
        sheet.getRange(i + 1, 6).setValue(JSON.stringify(completedItems));
        
        // Progress 셀에 명시적으로 텍스트 형식 설정
        const progressCell = sheet.getRange(i + 1, 8);
        progressCell.setNumberFormat('@');  // 텍스트 형식으로 설정
        progressCell.setValue(progress);
       
        // 모든 항목 완료 시 진행중검사에서 제거
        if (completedItems.length === totalItems.length) {
          sheet.deleteRow(i + 1);
          
          // 최종 결과 업데이트
          updateFinalResult(powderName, lotNumber);
        }
        
        break;
      }
    }
  } catch (error) {
    Logger.log('updateInspectionProgress 오류: ' + error.toString());
  }
}

// 최종 결과 업데이트
function updateFinalResult(powderName, lotNumber) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('InspectionResult');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === powderName && String(data[i][1]) === String(lotNumber)) {
        // 모든 항목의 결과 확인
        let finalResult = 'PASS';
        for (let j = 0; j < headers.length; j++) {
          if (headers[j].endsWith('Result') && headers[j] !== 'FinalResult') {
            if (data[i][j] === 'FAIL') {
              finalResult = 'FAIL';
              break;
            }
          }
        }
        
        const finalResultCol = headers.indexOf('FinalResult') + 1;
        if (finalResultCol > 0) {
          sheet.getRange(i + 1, finalResultCol).setValue(finalResult);
        }
        
        break;
      }
    }
  } catch (error) {
    Logger.log('updateFinalResult 오류: ' + error.toString());
  }
}

// 검사 결과 조회
function searchInspectionResults(powderName, lotNumber, startDate, endDate) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('InspectionResult');
    if (!sheet) {
      return { success: false, message: '검사 결과 시트를 찾을 수 없습니다.' };
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: true, data: [] }; // 데이터 없음
    }

    const headers = data[0];
    const results = [];

    for (let i = 1; i < data.length; i++) {
      let match = true;

      // 분말명 필터
      if (powderName && String(data[i][0]) !== String(powderName)) match = false;

      // LOT번호 필터
      if (lotNumber && String(data[i][1]) !== String(lotNumber)) match = false;

      // 날짜 필터 (InspectionTime = data[i][3])
      let inspectionDate = null;
      if (data[i][3]) {
        inspectionDate = new Date(data[i][3]);
        if (startDate && inspectionDate < new Date(startDate)) match = false;
        if (endDate && inspectionDate > new Date(endDate)) match = false;
      }

      if (match) {
        const finalResultIndex = headers.indexOf('FinalResult');
        const finalResult = data[i][finalResultIndex];
  
        // 완료된 검사만 포함 (FinalResult가 PASS 또는 FAIL인 경우)
        if (!finalResult || (finalResult !== 'PASS' && finalResult !== 'FAIL')) {
          continue; // 미완성 검사는 스킵
        }
        
        const result = {};
        for (let j = 0; j < headers.length; j++) {
          let value = data[i][j];

          // ✅ InspectionTime(4번째 열) 포맷팅 (yyyy-MM-dd HH:mm:ss)
          if (headers[j] === 'InspectionTime' && value) {
            value = Utilities.formatDate(new Date(value), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
          }

          result[headers[j]] = value;
        }
        results.push(result);
      }
    }

    // ✅ 검사시간 기준으로 내림차순 정렬 (최신순)
    results.sort(function(a, b) {
      const dateA = new Date(a.InspectionTime);
      const dateB = new Date(b.InspectionTime);
      return dateB - dateA; // 최신 날짜가 먼저 오도록
    });

    return { success: true, data: results };

  } catch (error) {
    Logger.log('searchInspectionResults 오류: ' + error.toString());
    return {
      success: false,
      message: '조회 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}

// 분말 목록 조회
function getPowderList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('PowderSpec');
    
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const powders = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        powders.push(data[i][0]);
      }
    }
    
    return powders;
    
  } catch (error) {
    Logger.log('getPowderList 오류: ' + error.toString());
    return [];
  }
}

// 디버그용 함수
function debugInspection() {
  Logger.log('=== 디버그 시작 ===');
  
  // 시트 존재 여부 확인
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('스프레드시트 ID: ' + ss.getId());
  
  const sheets = ['PowderSpec', 'InspectionResult', 'InspectionProgress'];
  sheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      Logger.log(sheetName + ' 시트 존재 - 행 수: ' + sheet.getLastRow());
    } else {
      Logger.log(sheetName + ' 시트 없음');
    }
  });
  
  // 분말 목록 확인
  const powders = getPowderList();
  Logger.log('분말 목록: ' + JSON.stringify(powders));
  
  // 테스트 검사 시작
  if (powders.length > 0) {
    const testResult = startInspection(powders[0], 'TEST001', '일상점검', '테스트');
    Logger.log('테스트 검사 시작 결과: ' + JSON.stringify(testResult));
  }
  
  Logger.log('=== 디버그 완료 ===');
}

// 입도분석 규격 조회
function getParticleSizeSpec(powderName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ParticleSize');
    
    if (!sheet) {
      Logger.log('ParticleSize 시트를 찾을 수 없습니다.');
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    const meshSpecs = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === powderName) {
        meshSpecs.push({
          meshSize: data[i][1],
          min: data[i][2],
          max: data[i][3]
        });
      }
    }
    
    return meshSpecs;
  } catch (error) {
    Logger.log('getParticleSizeSpec 오류: ' + error.toString());
    return [];
  }
}

// 입도분석 데이터 저장
function saveParticleSizeData(powderName, lotNumber, particleData) {
  try {
    Logger.log(`입도분석 저장 요청: ${powderName}, ${lotNumber}`);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const resultSheet = ss.getSheetByName('InspectionResult');
    const existingRowIndex = findExistingResultRow(powderName, lotNumber);
    
    // 전체 입도분석 결과 판정
    let overallResult = 'PASS';
    Object.values(particleData).forEach(meshData => {
      if (meshData.result === '불합격') {
        overallResult = 'FAIL';
      }
    });
    
    if (existingRowIndex > 0) {
      // 기존 행 업데이트
      updateParticleSizeResult(existingRowIndex, particleData, overallResult);
    } else {
      // 새 행 추가
      createNewParticleSizeResult(powderName, lotNumber, particleData, overallResult);
    }
    
    // 진행중검사 시트 업데이트
    updateInspectionProgress(powderName, lotNumber, 'ParticleSize');
    
    Logger.log('입도분석 저장 완료');
    
    return {
      success: true,
      result: overallResult
    };
    
  } catch (error) {
    Logger.log('saveParticleSizeData 오류: ' + error.toString());
    return {
      success: false,
      message: '저장 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}

// 입도분석 결과 업데이트
function updateParticleSizeResult(rowIndex, particleData, overallResult) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('InspectionResult');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // 각 MESH별 데이터 저장
    const meshMapping = {
      '80': 'ParticleSize80',
      '100': 'ParticleSize100', 
      '150': 'ParticleSize150',
      '200': 'ParticleSize200',
      '325': 'ParticleSize325',
      '325M': 'ParticleSize325M'
    };
    
    Object.keys(meshMapping).forEach(meshId => {
      const prefix = meshMapping[meshId];
      const data = particleData[meshId];
      
      if (data) {
        const val1Col = headers.indexOf(prefix + '_1') + 1;
        const val2Col = headers.indexOf(prefix + '_2') + 1;
        const avgCol = headers.indexOf(prefix + '_Avg') + 1;
        const resultCol = headers.indexOf(prefix + '_Result') + 1;
        
        if (val1Col > 0) sheet.getRange(rowIndex, val1Col).setValue(data.val1 || '');
        if (val2Col > 0) sheet.getRange(rowIndex, val2Col).setValue(data.val2 || '');
        if (avgCol > 0) sheet.getRange(rowIndex, avgCol).setValue(data.avg || '');
        if (resultCol > 0) sheet.getRange(rowIndex, resultCol).setValue(data.result === '합격' ? 'PASS' : 'FAIL');
      }
    });
    
    // 전체 입도분석 결과
    const overallCol = headers.indexOf('ParticleSizeResult') + 1;
    if (overallCol > 0) {
      sheet.getRange(rowIndex, overallCol).setValue(overallResult);
    }
    
  } catch (error) {
    Logger.log('updateParticleSizeResult 오류: ' + error.toString());
  }
}

// 새 입도분석 결과 행 생성  
function createNewParticleSizeResult(powderName, lotNumber, particleData, overallResult) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('InspectionResult');
    const existingInspection = getExistingInspection(powderName, lotNumber);
    
    // 새 행 데이터 준비
    const newRow = Array(sheet.getLastColumn()).fill('');
    newRow[0] = powderName;
    newRow[1] = lotNumber;
    newRow[2] = existingInspection.inspector;
    newRow[3] = new Date();
    newRow[4] = existingInspection.inspectionType;
    
    sheet.appendRow(newRow);
    
    // 입도분석 데이터 업데이트
    const newRowIndex = sheet.getLastRow();
    updateParticleSizeResult(newRowIndex, particleData, overallResult);
    
  } catch (error) {
    Logger.log('createNewParticleSizeResult 오류: ' + error.toString());
  }
}

// getInspectionDetail 함수에 분말 사양 정보 추가
function getInspectionDetail(powderName, lotNumber) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('InspectionResult');
    
    if (!sheet) {
      return { success: false, message: '검사 결과 시트를 찾을 수 없습니다.' };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === powderName && String(data[i][1]) === String(lotNumber)) {
        const result = {};
        
        // 모든 컬럼 데이터 매핑
        for (let j = 0; j < headers.length; j++) {
          let value = data[i][j];
          
          // 날짜 포맷팅
          if (headers[j] === 'InspectionTime' && value) {
            value = Utilities.formatDate(new Date(value), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
          }
          
          result[headers[j]] = value;
        }
        
        // 분말 사양 정보 추가
        const powderSpec = getPowderSpec(powderName);
        if (powderSpec) {
          result.powderSpec = powderSpec;
        }
        
        return { success: true, data: result };
      }
    }
    
    return { success: false, message: '해당 검사 데이터를 찾을 수 없습니다.' };
    
  } catch (error) {
    Logger.log('getInspectionDetail 오류: ' + error.toString());
    return {
      success: false,
      message: '조회 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}

// 겉보기밀도 저장 함수 (수정)
function saveApparentDensityItem(powderName, lotNumber, values) {
  try {
    // values 배열: [emptyCup1, powderWeight1, emptyCup2, powderWeight2, emptyCup3, powderWeight3]
    const apparentDensities = [];
    
    // 1차, 2차, 3차 각각 계산
    for (let i = 0; i < 3; i++) {
      const emptyCup = parseFloat(values[i * 2]);
      const powderWeight = parseFloat(values[i * 2 + 1]);
      
      if (!isNaN(emptyCup) && !isNaN(powderWeight) && emptyCup !== '' && powderWeight !== '') {
        const apparentDensity = (powderWeight - emptyCup) / 25;
        apparentDensities.push(apparentDensity);
      }
    }
    
    if (apparentDensities.length === 0) {
      return { success: false, message: '유효한 측정값이 없습니다.' };
    }
    
    // 평균 계산
    const average = apparentDensities.reduce((sum, val) => sum + val, 0) / apparentDensities.length;
    
    // 기존 검사 확인
    const existingInspection = getExistingInspection(powderName, lotNumber);
    if (!existingInspection) {
      return { success: false, message: '진행중인 검사를 찾을 수 없습니다.' };
    }
    
    // 규격 확인 및 합격/불합격 판정
    const items = getInspectionItems(powderName, existingInspection.inspectionType);
    const currentItem = items.find(item => item.name === 'ApparentDensity');
    
    let result = 'PASS';
    if (currentItem) {
      if (currentItem.min !== '' && average < parseFloat(currentItem.min)) {
        result = 'FAIL';
      }
      if (currentItem.max !== '' && average > parseFloat(currentItem.max)) {
        result = 'FAIL';
      }
    }
    
    // InspectionResult 시트에 저장
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const resultSheet = ss.getSheetByName('InspectionResult');
    const existingRowIndex = findExistingResultRow(powderName, lotNumber);
    
    // 저장할 값들: [emptyCup1, powderWeight1, emptyCup2, powderWeight2, emptyCup3, powderWeight3, average]
    const saveValues = [...values, average];
    
    if (existingRowIndex > 0) {
      updateInspectionResult(existingRowIndex, 'ApparentDensity', saveValues, average, result);
    } else {
      createNewInspectionResult(powderName, lotNumber, 'ApparentDensity', saveValues, average, result);
    }
    
    // 진행중검사 시트 업데이트
    updateInspectionProgress(powderName, lotNumber, 'ApparentDensity');
    
    return {
      success: true,
      average: average.toFixed(3),
      result: result
    };
    
  } catch (error) {
    Logger.log('saveApparentDensityItem 오류: ' + error.toString());
    return {
      success: false,
      message: '저장 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}

// 검사 결과 삭제 함수
function deleteInspectionResult(powderName, lotNumber) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const resultSheet = ss.getSheetByName('InspectionResult');
    
    if (!resultSheet) {
      return { success: false, message: '검사 결과 시트를 찾을 수 없습니다.' };
    }
    
    const data = resultSheet.getDataRange().getValues();
    
    // 해당 행 찾기
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === powderName && String(data[i][1]) === String(lotNumber)) {
        // 행 삭제
        resultSheet.deleteRow(i + 1);
        
        // 진행중검사 시트에서도 삭제
        const progressSheet = ss.getSheetByName('InspectionProgress');
        if (progressSheet) {
          const progressData = progressSheet.getDataRange().getValues();
          for (let j = 1; j < progressData.length; j++) {
            if (progressData[j][0] === powderName && String(progressData[j][1]) === String(lotNumber)) {
              progressSheet.deleteRow(j + 1);
              break;
            }
          }
        }
        
        Logger.log('검사 결과 삭제 완료: ' + powderName + ' / ' + lotNumber);
        return { success: true };
      }
    }
    
    return { success: false, message: '해당 검사 결과를 찾을 수 없습니다.' };
    
  } catch (error) {
    Logger.log('deleteInspectionResult 오류: ' + error.toString());
    return {
      success: false,
      message: '삭제 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}
