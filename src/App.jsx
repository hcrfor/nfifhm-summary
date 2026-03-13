import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import { Upload, Download, FileSpreadsheet, AlertCircle, CheckCircle2 } from 'lucide-react';
import _ from 'lodash';

function App() {
    const [data1, setData1] = useState([]); // 출현종 요약
    const [dataSummary, setDataSummary] = useState([]); // 출현종 요약 상단 '요약' 섹션
    const [data2, setData2] = useState([]); // 대경목 출현 요약
    const [dataMonitoring, setDataMonitoring] = useState([]); // 모니터링 요약
    const [loading, setLoading] = useState(false);
    const [error, setError] = useState(null);
    const [fileName, setFileName] = useState('');
    const [activeTab, setActiveTab] = useState('monitoring'); // monitoring | species | largeTrees
    const [fileType, setFileType] = useState('NFI'); // NFI | FHM
    const fileInputRef = useRef(null);
    const fhmFileInputRef = useRef(null);

    const handleFileUpload = (e, type) => {
        const file = e.target.files[0];
        if (!file) return;
        processFile(file, type);
    };

    const onDragOver = (e) => {
        e.preventDefault();
        e.currentTarget.classList.add('dragging');
    };

    const onDragLeave = (e) => {
        e.currentTarget.classList.remove('dragging');
    };

    const onDrop = (e, type) => {
        e.preventDefault();
        e.currentTarget.classList.remove('dragging');
        const file = e.dataTransfer.files[0];
        if (file) processFile(file, type);
    };

    const processFile = (file, type) => {
        setLoading(true);
        setError(null);
        setFileName(file.name);
        setFileType(type);

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const ab = e.target.result;
                const wb = XLSX.read(ab, { type: 'array' });

                // 공통 헬퍼 함수들 (강화된 버전)
                const clean = (str) => String(str || '').replace(/[^a-zA-Z0-9가-힣]/g, '');

                const getCol = (row, patterns, strict = false) => {
                    const keys = Object.keys(row);
                    // 1. 우선 순위 패턴 매칭 (정확도 향상)
                    for (const pattern of patterns) {
                        const cleanP = clean(pattern);
                        const foundKey = keys.find(key => {
                            const cleanK = clean(key);
                            if (strict) return cleanK === cleanP;
                            return cleanK.includes(cleanP);
                        });
                        if (foundKey) return row[foundKey];
                    }
                    return '';
                };

                const normalizeId = (id) => String(id || '').replace(/[^0-9]/g, '').trim();

                const readSheetData = (sheetKeywords, headerKeywords) => {
                    const actualSheetName = wb.SheetNames.find(name => {
                        const cleanName = clean(name);
                        return sheetKeywords.some(k => cleanName.includes(clean(k)));
                    });
                    if (!actualSheetName) return [];

                    const ws = wb.Sheets[actualSheetName];
                    const allRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
                    let headerIdx = -1;
                    
                    // 헤더 감지 로직 강화
                    for (let i = 0; i < Math.min(allRows.length, 20); i++) {
                        if (!allRows[i]) continue;
                        const rowStr = allRows[i].map(c => clean(c)).join('|');
                        // 헤더 키워드 중 2개 이상이 발견되면 헤더 행으로 간주 (정확도 확보)
                        const matchCount = headerKeywords.filter(k => rowStr.includes(clean(k))).length;
                        if (matchCount >= 1) {
                            headerIdx = i;
                            break;
                        }
                    }
                    if (headerIdx === -1) return [];
                    return XLSX.utils.sheet_to_json(ws, { range: headerIdx, raw: false, defval: '' });
                };

                // 타입별 시트/헤더 키워드 설정 (보강)
                const isFHM = type === 'FHM';
                const treeSheetKeywords = ['임목조사표', '임목조사', 'Tree'];
                const generalSheetKeywords = ['일반정보', '기본정보', 'General'];
                const standSheetKeywords = ['임분조사표', '임분조사', 'Stand'];

                // 1. 임목조사표 읽기
                const rawTreeJson = readSheetData(treeSheetKeywords, ['표본점', '수종']);
                if (rawTreeJson.length === 0) {
                    throw new Error(`${isFHM ? 'FHM' : 'NFI'} 양식에 맞는 임목조사표 시트를 찾을 수 없습니다.`);
                }

                let lastPointId = '';
                const treeProcessed = [];
                rawTreeJson.forEach(row => {
                    let rawPid = getCol(row, ['표본점번호', '표본점'], true);
                    if (!rawPid) rawPid = getCol(row, ['조사구', '번호']);

                    const speciesValue = String(getCol(row, ['수종명', '수종', '종명', '나무명']) || '').trim();
                    const headerKeywordsList = ['표본점', '수종명', '수종', '흉고'];
                    const isHeaderRow = (rawPid && headerKeywordsList.some(k => String(rawPid).includes(k))) || 
                                      (speciesValue && (headerKeywordsList.some(k => speciesValue.includes(k)) || speciesValue === 'undefined' || speciesValue === ''));

                    if (isHeaderRow) return;

                    let normalizedPid = normalizeId(rawPid);
                    let currentPid = '';

                    if (!normalizedPid || normalizedPid === 'undefined' || normalizedPid.length < 5) {
                        currentPid = lastPointId;
                    } else {
                        currentPid = normalizedPid;
                        lastPointId = normalizedPid;
                    }

                    if (currentPid && speciesValue && speciesValue !== 'undefined' && speciesValue !== '') {
                        treeProcessed.push({
                            pointId: currentPid,
                            species: speciesValue,
                            height: getCol(row, ['수고', '나무수고']),
                            dbh: getCol(row, ['흉고직경', '직경', 'DBH']),
                            dist: getCol(row, ['거리']),
                            azimuth: getCol(row, ['방위']),
                            note: String(getCol(row, ['비고', '코드'])).replace('undefined', '').trim()
                        });
                    }
                });

                // 2. 일반정보 읽기 (토지이용정보)
                const rawGeneralJson = readSheetData(generalSheetKeywords, ['표본점', '토지이용']);
                const generalMap = {};
                const landUseCodes = {
                    '1': '임목지', '2': '미립목지', '3': '제지', '4': '경작지',
                    '5': '초지', '6': '습지', '7': '주거지', '8': '기타', '95': '죽림'
                };
                rawGeneralJson.forEach(row => {
                    const pid = normalizeId(getCol(row, ['표본점번호', '표본점']));
                    const code = String(getCol(row, ['토지이용정보', '토지이용'])).trim();
                    if (pid && pid.length >= 5) {
                        generalMap[pid] = landUseCodes[code] || code;
                    }
                });

                // 3. 임분조사표 읽기
                const rawStandJson = readSheetData(standSheetKeywords, ['표본점', '임종', '임상']);
                const standMap = {};
                const forestClassCodes = { '0': '천연림', '1': '인공림' };
                const regenCodes = { '0': '기타', '1': '조림', '2': '천연하종', '3': '맹아' };
                const forestTypeCodes = { '0': '침엽수림', '1': '활엽수림', '2': '혼효림', '3': '비산림' };

                rawStandJson.forEach(row => {
                    const pid = normalizeId(getCol(row, ['표본점번호', '표본점']));
                    if (pid && pid.length >= 5) {
                        const fclass = String(getCol(row, ['임종', 'FCLAS'])).trim();
                        const regen = String(getCol(row, ['갱신형태', 'REGEN'])).trim();
                        const ftype = String(getCol(row, ['임상', 'FTYPE'])).trim();

                        standMap[pid] = {
                            fclass: forestClassCodes[fclass] || fclass,
                            regen: regenCodes[regen] || regen,
                            ftype: forestTypeCodes[ftype] || ftype,
                            dclass: getCol(row, ['경급', 'DCLAS']),
                            aclas: getCol(row, ['영급', 'ACLAS']),
                            nonForestBasic: getCol(row, ['비산림면적기본', '비산림']),
                            nonForestLarge: getCol(row, ['비산림면적대경목', '대경목비산림'])
                        };
                    }
                });

                // --- 데이터 요약 생성 ---
                const speciesSummary = [];
                const topWinnerSummary = [];
                const monitoringSummary = [];
                const groupedByPoint = _.groupBy(treeProcessed, 'pointId');
                const allPointIdsFound = new Set();
                treeProcessed.forEach(t => allPointIdsFound.add(t.pointId));
                Object.keys(generalMap).forEach(p => allPointIdsFound.add(p));
                Object.keys(standMap).forEach(p => allPointIdsFound.add(p));

                const expandedPointsSet = new Set();
                allPointIdsFound.forEach(pid => {
                    const sPid = String(pid).trim();
                    if (!sPid || sPid.length < 5) return;
                    const lastChar = sPid.slice(-1);
                    const base = (['1', '2', '3', '4'].includes(lastChar)) ? sPid.slice(0, -1) : sPid;
                    if (base) {
                        expandedPointsSet.add(base + '1');
                        expandedPointsSet.add(base + '2');
                        expandedPointsSet.add(base + '3');
                        expandedPointsSet.add(base + '4');
                    }
                });
                const summaryPoints = Array.from(expandedPointsSet).sort();

                summaryPoints.forEach(pointId => {
                    const pointData = groupedByPoint[pointId] || [];
                    const groupedBySpecies = _.groupBy(pointData, 'species');
                    const sortedSpeciesNames = Object.keys(groupedBySpecies).sort((a, b) => a.localeCompare(b));

                    let pCount = 0;
                    let pHeights = [];
                    let pWinnerSpeciesList = [];
                    let pMaxH = -1;

                    sortedSpeciesNames.forEach(speciesName => {
                        const rows = groupedBySpecies[speciesName];
                        const hs = rows.map(r => parseFloat(r.height)).filter(h => !isNaN(h));
                        pCount += rows.length;
                        pHeights.push(...hs);
                        const mx = hs.length > 0 ? _.max(hs) : null;
                        if (mx !== null) {
                            if (mx > pMaxH) { pMaxH = mx; pWinnerSpeciesList = [speciesName]; }
                            else if (mx === pMaxH) { pWinnerSpeciesList.push(speciesName); }
                        }
                    });

                    const tMaxH = pHeights.length > 0 ? _.max(pHeights) : null;
                    const tAvgH = pHeights.length > 0 ? _.mean(pHeights) : null;
                    const sData = standMap[pointId] || {};
                    
                    monitoringSummary.push({
                        pointId: pointId, landUse: generalMap[pointId] || '', fclass: sData.fclass || '',
                        regen: sData.regen || '', ftype: sData.ftype || '', dclass: sData.dclass || '', aclas: sData.aclas || '',
                        totalStems: pCount, maxHSpecies: pWinnerSpeciesList.join(', '),
                        maxH: tMaxH !== null ? Math.round(tMaxH) : '', avgH: tAvgH !== null ? Math.round(tAvgH) : '',
                        nonForestBasic: sData.nonForestBasic || '0', nonForestLarge: sData.nonForestLarge || '0'
                    });

                    let pointSpeciesList = [];
                    sortedSpeciesNames.forEach(speciesName => {
                        const speciesRows = groupedBySpecies[speciesName];
                        pointSpeciesList.push({
                            type: 'data', label: speciesName, pointId: pointId, count: speciesRows.length,
                            winnerSpecies: '', maxHeight: '', avgHeight: ''
                        });
                    });

                    const pTotalMaxH = pHeights.length > 0 ? _.max(pHeights) : null;
                    const pTotalAvgH = pHeights.length > 0 ? _.mean(pHeights) : null;
                    const subtotalRow = {
                        type: 'subtotal', label: '소계', pointId: pointId, count: pCount,
                        winnerSpecies: pWinnerSpeciesList.join(', '),
                        maxHeight: pTotalMaxH !== null ? Math.round(pTotalMaxH) : '',
                        avgHeight: pTotalAvgH !== null ? Math.round(pTotalAvgH) : ''
                    };

                    speciesSummary.push({ type: 'header', label: pointId, pointId: pointId, count: '', winnerSpecies: '', maxHeight: '', avgHeight: '' });
                    speciesSummary.push(subtotalRow);
                    speciesSummary.push(...pointSpeciesList);
                    topWinnerSummary.push({ label: pointId, count: subtotalRow.count, winnerSpecies: subtotalRow.winnerSpecies, maxHeight: subtotalRow.maxHeight, avgHeight: subtotalRow.avgHeight });
                });

                setData1(speciesSummary);
                setDataSummary(topWinnerSummary);
                setDataMonitoring(monitoringSummary);

                const largeTrees = treeProcessed.filter(item => {
                    const dbh = parseFloat(item.dbh);
                    const note = String(item.note).toUpperCase();
                    return !isNaN(dbh) && dbh >= 30 && (note === '' || note.includes('L'));
                });

                const largeTreesByPoint = _.groupBy(largeTrees, 'pointId');
                const sortedLargeTrees = [];
                summaryPoints.forEach(pointId => {
                    const treeList = largeTreesByPoint[pointId] || [];
                    if (treeList.length > 0) {
                        _.orderBy(treeList, ['species', (item) => parseFloat(item.dbh)], ['asc', 'desc']).forEach(item => {
                            sortedLargeTrees.push({ pointId: item.pointId, species: item.species, dbh: item.dbh, combined: `${item.species}${item.dbh}`, dist: item.dist, azimuth: item.azimuth, note: item.note });
                        });
                    } else {
                        sortedLargeTrees.push({ pointId: pointId, species: '', dbh: '', combined: '', dist: '', azimuth: '', note: '' });
                    }
                });

                setData2(sortedLargeTrees);
                setLoading(false);
            } catch (err) {
                console.error(err);
                setError(err.message || '파일 처리 중 오류가 발생했습니다.');
                setLoading(false);
            }
        };
        reader.readAsArrayBuffer(file);
    };

    const downloadExcel = () => {
        const wb = XLSX.utils.book_new();

        // 1. 모니터링 요약
        const wsMonData = [
            ['표본점번호', '토지이용', '임종', '갱신형태', '임상', '경급', '영급', '총본수', '최대 수고 수종명', '최대 수고', '평균 수고', '비산림면적(기본조사원)', '비산림면적(대경목조사원)']
        ];
        dataMonitoring.forEach(row => {
            wsMonData.push([
                row.pointId, row.landUse, row.fclass, row.regen, row.ftype, row.dclass, row.aclas,
                row.totalStems, row.maxHSpecies, row.maxH, row.avgH, row.nonForestBasic, row.nonForestLarge
            ]);
        });
        const wsMon = XLSX.utils.aoa_to_sheet(wsMonData);
        XLSX.utils.book_append_sheet(wb, wsMon, '모니터링 요약');

        // 2. 출현종 요약
        const ws1Data = [
            ['요약', '', '', '', ''],
            ['레이블', '개수', '수종명', '수고 최대값', '평균값']
        ];
        dataSummary.forEach(row => {
            ws1Data.push([row.label, row.count, row.winnerSpecies, row.maxHeight, row.avgHeight]);
        });
        ws1Data.push(['', '', '', '', '']);
        ws1Data.push(['', '', '', '', '']);
        ws1Data.push(['레이블', '개수', '수종명', '수고 최대값', '평균값']);
        data1.forEach(row => {
            ws1Data.push([row.label, row.count, row.winnerSpecies, row.maxHeight, row.avgHeight]);
        });
        const ws1 = XLSX.utils.aoa_to_sheet(ws1Data);
        XLSX.utils.book_append_sheet(wb, ws1, '출현종 요약');

        // 3. 대경목 출현 요약
        const ws2Data = [
            ['표본점번호', '수종명', '흉고직경', '수종명 흉고직경', '거리', '방위', '비고']
        ];
        data2.forEach(row => {
            ws2Data.push([row.pointId, row.species, row.dbh, row.combined, row.dist, row.azimuth, row.note]);
        });
        const ws2 = XLSX.utils.aoa_to_sheet(ws2Data);
        XLSX.utils.book_append_sheet(wb, ws2, '대경목 출현 요약');

        // 파일명 생성 로직: 첫 번째 표본점 번호의 마지막 자리를 제외하고 '_모니터링 요약'을 붙임
        const firstPointId = dataMonitoring.length > 0 ? String(dataMonitoring[0].pointId).trim() : '';
        const fileNamePrefix = firstPointId ? firstPointId.slice(0, -1) : '임목조사';
        const finalFileName = `${fileNamePrefix}_모니터링 요약.xlsx`;

        XLSX.writeFile(wb, finalFileName);
    };

    return (
        <div className="dashboard">
            <header>
                <h1>임목조사 데이터 요약 도구</h1>
                <p>숲의 데이터를 정확하고 빠르게 분석합니다</p>
            </header>

            {error && (
                <div className="error-msg">
                    <AlertCircle size={20} />
                    <span>{error}</span>
                </div>
            )}

            {!data1.length ? (
                <div className="flex flex-col md:flex-row gap-6 mb-8">
                    {/* NFI 업로드 구역 */}
                    <div
                        className="upload-area flex-1"
                        onDragOver={onDragOver}
                        onDragLeave={onDragLeave}
                        onDrop={(e) => onDrop(e, 'NFI')}
                        onClick={() => fileInputRef.current.click()}
                    >
                        <div className="type-badge nfi">NFI 전용</div>
                        <Upload className="upload-icon mx-auto" strokeWidth={1.5} />
                        <p className="text-xl font-semibold mb-2">NFI 파일 업로드</p>
                        <p className="text-gray-500 text-sm">국가산림자원조사 파일을 선택하세요</p>
                        <input
                            type="file"
                            className="hidden"
                            ref={fileInputRef}
                            onChange={(e) => handleFileUpload(e, 'NFI')}
                            accept=".xlsx, .xlsm"
                        />
                    </div>

                    {/* FHM 업로드 구역 */}
                    <div
                        className="upload-area flex-1 border-emerald-200 hover:border-emerald-500 hover:bg-emerald-50/30"
                        onDragOver={onDragOver}
                        onDragLeave={onDragLeave}
                        onDrop={(e) => onDrop(e, 'FHM')}
                        onClick={() => fhmFileInputRef.current.click()}
                    >
                        <div className="type-badge fhm">FHM 전용</div>
                        <Upload className="upload-icon mx-auto text-emerald-500" strokeWidth={1.5} />
                        <p className="text-xl font-semibold mb-2">FHM 파일 업로드</p>
                        <p className="text-gray-500 text-sm">산림건강성모니터링 파일을 선택하세요</p>
                        <input
                            type="file"
                            className="hidden"
                            ref={fhmFileInputRef}
                            onChange={(e) => handleFileUpload(e, 'FHM')}
                            accept=".xlsx, .xlsm"
                        />
                    </div>
                </div>
            ) : (
                <div className="results-section">
                    <div className="flex justify-between items-center mb-6">
                        <div className="flex items-center gap-3">
                            <div className="bg-blue-100 p-2 rounded-lg">
                                <FileSpreadsheet className="text-blue-600" size={24} />
                            </div>
                            <div>
                                <div className="flex items-center gap-2">
                                    <span className={`px-2 py-0.5 rounded text-[10px] font-bold uppercase ${fileType === 'FHM' ? 'bg-emerald-100 text-emerald-700' : 'bg-blue-100 text-blue-700'}`}>
                                        {fileType}
                                    </span>
                                    <p className="font-bold text-lg">{fileName}</p>
                                </div>
                                <div className="flex items-center gap-1 text-green-600 text-sm">
                                    <CheckCircle2 size={14} />
                                    <span>분석 완료</span>
                                </div>
                            </div>
                        </div>
                        <button
                            className="btn btn-primary"
                            onClick={downloadExcel}
                        >
                            <Download size={18} />
                            요약 파일 다운로드
                        </button>
                    </div>

                    <div className="tabs">
                        <div
                            className={`tab ${activeTab === 'monitoring' ? 'active' : ''}`}
                            onClick={() => setActiveTab('monitoring')}
                        >
                            모니터링 요약
                        </div>
                        <div
                            className={`tab ${activeTab === 'species' ? 'active' : ''}`}
                            onClick={() => setActiveTab('species')}
                        >
                            출현종 요약
                        </div>
                        <div
                            className={`tab ${activeTab === 'largeTrees' ? 'active' : ''}`}
                            onClick={() => setActiveTab('largeTrees')}
                        >
                            대경목 출현 요약
                        </div>
                    </div>

                    <div className={`table-container ${activeTab !== 'monitoring' ? 'hidden' : ''}`}>
                        <table>
                            <thead>
                                <tr>
                                    <th>표본점번호</th>
                                    <th>토지이용</th>
                                    <th>임종</th>
                                    <th>갱신형태</th>
                                    <th>임상</th>
                                    <th>경급</th>
                                    <th>영급</th>
                                    <th>총본수</th>
                                    <th>최대 수고 수종명</th>
                                    <th>최대 수고</th>
                                    <th>평균 수고</th>
                                    <th>비산림면적(기본)</th>
                                    <th>비산림면적(대경목)</th>
                                </tr>
                            </thead>
                            <tbody>
                                {dataMonitoring.map((row, idx) => (
                                    <tr key={idx}>
                                        <td>{row.pointId}</td>
                                        <td>{row.landUse}</td>
                                        <td>{row.fclass}</td>
                                        <td>{row.regen}</td>
                                        <td>{row.ftype}</td>
                                        <td>{row.dclass}</td>
                                        <td>{row.aclas}</td>
                                        <td>{row.totalStems}</td>
                                        <td>{row.maxHSpecies}</td>
                                        <td>{row.maxH}</td>
                                        <td>{row.avgH}</td>
                                        <td>{row.nonForestBasic}</td>
                                        <td>{row.nonForestLarge}</td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>

                    <div className={`table-container ${activeTab !== 'species' ? 'hidden' : ''}`}>
                        {/* 상단 요약 테이블 */}
                        <div className="mb-8 overflow-x-auto rounded-xl border border-blue-100 bg-blue-50/30 p-4">
                            <h3 className="mb-3 font-bold text-blue-800">단위 표본점별 요약</h3>
                            <table className="mb-0">
                                <thead>
                                    <tr>
                                        <th>레이블</th>
                                        <th>개수</th>
                                        <th>수종명</th>
                                        <th>수고 최대값</th>
                                        <th>평균값</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {dataSummary.map((row, idx) => (
                                        <tr key={idx} className="bg-white">
                                            <td>{row.label}</td>
                                            <td>{row.count}</td>
                                            <td>{row.winnerSpecies}</td>
                                            <td>{row.maxHeight}</td>
                                            <td>{row.avgHeight}</td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>

                        {/* 메인 상세 테이블 */}
                        <table>
                            <thead>
                                <tr>
                                    <th>레이블</th>
                                    <th>개수</th>
                                    <th>수종명</th>
                                    <th>수고 최대값</th>
                                    <th>평균값</th>
                                </tr>
                            </thead>
                            <tbody>
                                {data1.map((row, idx) => (
                                    <tr key={idx} className={row.type === 'subtotal' ? 'subtotal' : row.type === 'header' ? 'point-header' : ''}>
                                        <td style={{ paddingLeft: row.type === 'data' ? '2.5rem' : '1rem' }}>
                                            {row.label}
                                        </td>
                                        <td>{row.count}</td>
                                        <td>{row.winnerSpecies}</td>
                                        <td>{row.maxHeight}</td>
                                        <td>{row.avgHeight}</td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>

                    <div className={`table-container ${activeTab !== 'largeTrees' ? 'hidden' : ''}`}>
                        <table>
                            <thead>
                                <tr>
                                    <th>표본점번호</th>
                                    <th>수종명</th>
                                    <th>흉고직경</th>
                                    <th>수종명 흉고직경</th>
                                    <th>거리</th>
                                    <th>방위</th>
                                    <th>비고</th>
                                </tr>
                            </thead>
                            <tbody>
                                {data2.map((row, idx) => (
                                    <tr key={idx}>
                                        <td>{row.pointId}</td>
                                        <td>{row.species}</td>
                                        <td>{row.dbh}</td>
                                        <td>{row.combined}</td>
                                        <td>{row.dist}</td>
                                        <td>{row.azimuth}</td>
                                        <td>{row.note}</td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>

                    <div className="actions">
                        <button
                            className="text-gray-500 hover:text-gray-700 text-sm font-medium"
                            onClick={() => {
                                setData1([]);
                                setData2([]);
                                setDataMonitoring([]);
                                setFileName('');
                                setError(null);
                            }}
                        >
                            다른 파일 분석하기
                        </button>
                    </div>
                </div>
            )}

            {loading && (
                <div className="fixed inset-0 bg-black/20 backdrop-blur-sm flex items-center justify-center z-50">
                    <div className="bg-white p-6 rounded-2xl shadow-2xl flex flex-col items-center gap-4">
                        <div className="loading-spinner" style={{ borderTopColor: '#2563eb' }}></div>
                        <p className="font-semibold">데이터 분석 중...</p>
                    </div>
                </div>
            )}
        </div>
    );
}

export default App;
