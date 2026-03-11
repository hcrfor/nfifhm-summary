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
    const fileInputRef = useRef(null);

    const handleFileUpload = (e) => {
        const file = e.target.files[0];
        if (!file) return;
        processFile(file);
    };

    const onDragOver = (e) => {
        e.preventDefault();
        e.currentTarget.classList.add('dragging');
    };

    const onDragLeave = (e) => {
        e.currentTarget.classList.remove('dragging');
    };

    const onDrop = (e) => {
        e.preventDefault();
        e.currentTarget.classList.remove('dragging');
        const file = e.dataTransfer.files[0];
        if (file) processFile(file);
    };

    const processFile = (file) => {
        setLoading(true);
        setError(null);
        setFileName(file.name);

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const ab = e.target.result;
                const wb = XLSX.read(ab, { type: 'array' });

                // 공통 헬퍼 함수들
                const getCol = (row, patterns) => {
                    const foundKey = Object.keys(row).find(key =>
                        patterns.some(p => key.replace(/\s/g, '').includes(p))
                    );
                    return foundKey ? row[foundKey] : '';
                };

                const readSheetData = (sheetName, headerKeyword) => {
                    if (!wb.SheetNames.includes(sheetName)) return [];
                    const ws = wb.Sheets[sheetName];
                    const allRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
                    let headerIdx = 0;
                    for (let i = 0; i < Math.min(allRows.length, 10); i++) {
                        const row = allRows[i].map(c => String(c).replace(/\s/g, ''));
                        if (row.some(c => c.includes(headerKeyword))) {
                            headerIdx = i;
                            break;
                        }
                    }
                    return XLSX.utils.sheet_to_json(ws, { range: headerIdx, raw: false, defval: '' });
                };

                // 1. 임목조사표 읽기
                const rawTreeJson = readSheetData('임목조사표', '표본점번호');
                if (rawTreeJson.length === 0) {
                    throw new Error('임목조사표 시트를 확인 하시기 바랍니다.');
                }

                const treeProcessed = rawTreeJson.map(row => ({
                    pointId: String(getCol(row, ['표본점번호', '표본점'])),
                    species: String(getCol(row, ['수종명', '수종'])),
                    height: getCol(row, ['수고(cm)', '수고']),
                    dbh: getCol(row, ['흉고직경', '직경']),
                    dist: getCol(row, ['거리']),
                    azimuth: getCol(row, ['방위']),
                    note: String(getCol(row, ['비고(개체목구분코드)', '비고', '코드']))
                })).filter(item => item.pointId && item.pointId !== 'undefined' && item.pointId !== '');

                // 2. 일반정보 읽기 (토지이용정보)
                const rawGeneralJson = readSheetData('일반정보', '표본점번호');
                const generalMap = {};
                const landUseCodes = {
                    '1': '임목지', '2': '미립목지', '3': '제지', '4': '경작지',
                    '5': '초지', '6': '습지', '7': '주거지', '8': '기타', '95': '죽림'
                };
                rawGeneralJson.forEach(row => {
                    const pid = String(getCol(row, ['표본점번호', '표본점']));
                    const code = String(getCol(row, ['토지이용정보', '토지이용'])).trim();
                    if (pid) {
                        generalMap[pid] = landUseCodes[code] || code;
                    }
                });

                // 3. 임분조사표 읽기
                const rawStandJson = readSheetData('임분조사표', '표본점번호');
                const standMap = {};
                const forestClassCodes = { '0': '천연림', '1': '인공림' };
                const regenCodes = { '0': '기타', '1': '조림', '2': '천연하종', '3': '맹아' };
                const forestTypeCodes = { '0': '침엽수림', '1': '활엽수림', '2': '혼효림', '3': '비산림' };

                rawStandJson.forEach(row => {
                    const pid = String(getCol(row, ['표본점번호', '표본점']));
                    if (pid) {
                        const fclass = String(getCol(row, ['임종', 'FCLAS'])).trim();
                        const regen = String(getCol(row, ['갱신형태', 'REGEN'])).trim();
                        const ftype = String(getCol(row, ['임상', 'FTYPE'])).trim();

                        standMap[pid] = {
                            fclass: forestClassCodes[fclass] || fclass,
                            regen: regenCodes[regen] || regen,
                            ftype: forestTypeCodes[ftype] || ftype,
                            dclass: getCol(row, ['경급', 'DCLAS']),
                            aclas: getCol(row, ['영급', 'ACLAS']),
                            nonForestBasic: getCol(row, ['비산림면적기본조사원', '비산림면적(기본']),
                            nonForestLarge: getCol(row, ['비산림면적대경목조사원', '비산림면적(대경목'])
                        };
                    }
                });

                // --- Logic: 데이터 요약 생성 ---
                const speciesSummary = [];
                const topWinnerSummary = [];
                const monitoringSummary = [];

                const groupedByPoint = _.groupBy(treeProcessed, 'pointId');
                const sortedPoints = Object.keys(groupedByPoint).sort();

                sortedPoints.forEach(pointId => {
                    const pointData = groupedByPoint[pointId];
                    const groupedBySpecies = _.groupBy(pointData, 'species');
                    const sortedSpeciesNames = Object.keys(groupedBySpecies).sort((a, b) => a.localeCompare(b));

                    let pointCountTotal = 0;
                    let pointHeights = [];
                    let pointSpeciesList = [];

                    // 각 표본점별 모든 나무에서 최대 수고 정보 찾기
                    let pointMaxH = -1;
                    let pointMaxHName = '';

                    sortedSpeciesNames.forEach(speciesName => {
                        if (!speciesName || speciesName === 'undefined') return;

                        const speciesRows = groupedBySpecies[speciesName];
                        const heights = speciesRows
                            .map(r => parseFloat(r.height))
                            .filter(h => !isNaN(h));

                        const count = speciesRows.length;
                        const maxH = heights.length > 0 ? _.max(heights) : null;
                        const avgH = heights.length > 0 ? _.mean(heights) : null;

                        pointCountTotal += count;
                        pointHeights.push(...heights);

                        if (maxH !== null && maxH > pointMaxH) {
                            pointMaxH = maxH;
                            pointMaxHName = speciesName;
                        }

                        pointSpeciesList.push({
                            type: 'data',
                            label: speciesName,
                            pointId: pointId,
                            count: count,
                            winnerSpecies: '',
                            maxHeight: maxH !== null ? Math.round(maxH) : '',
                            avgHeight: avgH !== null ? Math.round(avgH) : ''
                        });
                    });

                    // Sheet 1 처리 (상세/요약)
                    const pTotalMaxH = pointHeights.length > 0 ? _.max(pointHeights) : null;
                    const pTotalAvgH = pointHeights.length > 0 ? _.mean(pointHeights) : null;

                    let winnerNames = '';
                    if (pTotalMaxH !== null) {
                        winnerNames = pointSpeciesList
                            .filter(s => s.maxHeight !== '' && Number(s.maxHeight) === Math.round(pTotalMaxH))
                            .map(s => s.label)
                            .join(', ');
                    }

                    const subtotalRow = {
                        type: 'subtotal',
                        label: '소계',
                        pointId: pointId,
                        count: pointCountTotal,
                        winnerSpecies: winnerNames,
                        maxHeight: pTotalMaxH !== null ? Math.round(pTotalMaxH) : '',
                        avgHeight: pTotalAvgH !== null ? Math.round(pTotalAvgH) : ''
                    };

                    speciesSummary.push({
                        type: 'header',
                        label: pointId,
                        pointId: pointId,
                        count: '', winnerSpecies: '', maxHeight: '', avgHeight: ''
                    });
                    speciesSummary.push(subtotalRow);
                    speciesSummary.push(...pointSpeciesList);

                    topWinnerSummary.push({
                        label: pointId,
                        count: subtotalRow.count,
                        winnerSpecies: subtotalRow.winnerSpecies,
                        maxHeight: subtotalRow.maxHeight,
                        avgHeight: subtotalRow.avgHeight
                    });

                    // 모니터링 요약 데이터 구성
                    const sData = standMap[pointId] || {};
                    monitoringSummary.push({
                        pointId: pointId,
                        landUse: generalMap[pointId] || '',
                        fclass: sData.fclass || '',
                        regen: sData.regen || '',
                        ftype: sData.ftype || '',
                        dclass: sData.dclass || '',
                        aclas: sData.aclas || '',
                        totalStems: pointCountTotal,
                        maxHSpecies: winnerNames, // 최대 수고 수종명들
                        maxH: pTotalMaxH !== null ? Math.round(pTotalMaxH) : '',
                        avgH: pTotalAvgH !== null ? Math.round(pTotalAvgH) : '',
                        nonForestBasic: sData.nonForestBasic || '0',
                        nonForestLarge: sData.nonForestLarge || '0'
                    });
                });

                setData1(speciesSummary);
                setDataSummary(topWinnerSummary);
                setDataMonitoring(monitoringSummary);

                // --- 대경목 logic (기존 유지) ---
                const largeTrees = treeProcessed.filter(item => {
                    const dbh = parseFloat(item.dbh);
                    const note = String(item.note).toUpperCase();
                    const isDbhOk = !isNaN(dbh) && dbh >= 30;
                    const isNoteOk = note === '' || note === 'undefined' || note.includes('L');
                    return isDbhOk && isNoteOk;
                });

                const sortedLargeTrees = _.orderBy(largeTrees,
                    ['pointId', 'species', (item) => parseFloat(item.dbh)],
                    ['asc', 'asc', 'desc']
                ).map(item => ({
                    pointId: item.pointId,
                    species: item.species,
                    dbh: item.dbh,
                    combined: `${item.species}${item.dbh}`,
                    dist: item.dist,
                    azimuth: item.azimuth,
                    note: (item.note === 'undefined' ? '' : item.note)
                }));

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
                <div
                    className="upload-area"
                    onDragOver={onDragOver}
                    onDragLeave={onDragLeave}
                    onDrop={onDrop}
                    onClick={() => fileInputRef.current.click()}
                >
                    <Upload className="upload-icon mx-auto" strokeWidth={1.5} />
                    <p className="text-xl font-semibold mb-2">엑셀 파일 업로드</p>
                    <p className="text-gray-500">파일을 선택하거나 여기로 드래그하세요 (.xlsx, .xlsm)</p>
                    <input
                        type="file"
                        className="hidden"
                        ref={fileInputRef}
                        onChange={handleFileUpload}
                        accept=".xlsx, .xlsm"
                    />
                </div>
            ) : (
                <div className="results-section">
                    <div className="flex justify-between items-center mb-6">
                        <div className="flex items-center gap-3">
                            <div className="bg-blue-100 p-2 rounded-lg">
                                <FileSpreadsheet className="text-blue-600" size={24} />
                            </div>
                            <div>
                                <p className="font-bold text-lg">{fileName}</p>
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
