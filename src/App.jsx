import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { Upload, Download, FileSpreadsheet, AlertCircle, CheckCircle2 } from 'lucide-react';
import _ from 'lodash';

function App() {
    const [data1, setData1] = useState([]); // 출현종 요약
    const [dataSummary, setDataSummary] = useState([]); // 출현종 요약 상단 '요약' 섹션
    const [data2, setData2] = useState([]); // 대경목 출현 요약
    const [dataMonitoring, setDataMonitoring] = useState([]); // 모니터링 요약
    const [data2021_1, setData2021_1] = useState([]); // 2021 출현종 요약
    const [data2021_Summary, setData2021_Summary] = useState([]); // 2021 출현종 요약 상단 '요약' 섹션
    const [data2021_2, setData2021_2] = useState([]); // 2021 대경목 출현 요약
    const [dataComparison, setDataComparison] = useState([]); // 모니터링비교_출현수종 추가
    const [loading, setLoading] = useState(false);
    const [error, setError] = useState(null);
    const [fileName, setFileName] = useState('');
    const [activeTab, setActiveTab] = useState('monitoring'); // monitoring | species | largeTrees | species2021 | largeTrees2021 | comparison
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

                // 엑셀 범위(!ref)가 잘못되어 데이터가 잘리는 문제 해결을 위한 함수
                const updateRange = (ws) => {
                    if (!ws) return;
                    const keys = Object.keys(ws).filter(k => k && k[0] !== '!');
                    if (keys.length === 0) return;
                    let minR = Infinity, maxR = -Infinity, minC = Infinity, maxC = -Infinity;
                    keys.forEach(key => {
                        try {
                            const cell = XLSX.utils.decode_cell(key);
                            if (cell.r < minR) minR = cell.r;
                            if (cell.r > maxR) maxR = cell.r;
                            if (cell.c < minC) minC = cell.c;
                            if (cell.c > maxC) maxC = cell.c;
                        } catch (e) {}
                    });
                    if (minR !== Infinity) {
                        ws['!ref'] = XLSX.utils.encode_range({ s: { r: minR, c: minC }, e: { r: maxR, c: maxC } });
                    }
                };

                const getCol = (row, patterns, strict = false) => {
                    const keys = Object.keys(row);
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

                const getClusterValue = (map, pid) => {
                    if (map[pid]) return map[pid];
                    const sPid = String(pid);
                    const lastChar = sPid.slice(-1);
                    const base = (['1', '2', '3', '4'].includes(lastChar)) ? sPid.slice(0, -1) : sPid;
                    return map[base] || map[base + '1'] || map[base + '2'] || map[base + '3'] || map[base + '4'] || '';
                };

                const getClusterData = (map, pid) => {
                    if (map[pid]) return map[pid];
                    const sPid = String(pid);
                    const lastChar = sPid.slice(-1);
                    const base = (['1', '2', '3', '4'].includes(lastChar)) ? sPid.slice(0, -1) : sPid;
                    return map[base] || map[base + '1'] || map[base + '2'] || map[base + '3'] || map[base + '4'] || {};
                };

                const generateSummaries = (trees, gMap = {}, sMap = {}, customPoints = null) => {
    const speciesSummary = [];
    const monitoringSummary = [];
    const groupedByPoint = _.groupBy(trees, 'pointId');
    
    let summaryPoints = [];
    if (customPoints) {
        summaryPoints = customPoints;
    } else {
        const allPointIdsFound = new Set();
        trees.forEach(t => allPointIdsFound.add(t.pointId));
        Object.keys(gMap).forEach(p => allPointIdsFound.add(p));
        Object.keys(sMap).forEach(p => allPointIdsFound.add(p));

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
        summaryPoints = Array.from(expandedPointsSet).sort();
    }

    summaryPoints.forEach(pointId => {
        const pointData = groupedByPoint[pointId] || [];
        const sData = getClusterData(sMap, pointId);

        if (pointData.length === 0 && !sData.fclass && !sData.ftype) return;

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

        const tAvgH = pHeights.length > 0 ? _.mean(pHeights) : null;
        
        monitoringSummary.push({
            pointId, 
            landUse: getClusterValue(gMap, pointId), 
            fclass: sData.fclass || '',
            regen: sData.regen || '', ftype: sData.ftype || '', dclass: sData.dclass || '', aclas: sData.aclas || '',
            totalStems: pCount, maxHSpecies: pWinnerSpeciesList.join(', '),
            avgH: tAvgH !== null ? Math.round(tAvgH) : '',
            nonForestBasic: sData.nonForestBasic || '0', nonForestLarge: sData.nonForestLarge || '0'
        });

        let pointSpeciesList = [];
        sortedSpeciesNames.forEach(speciesName => {
            const speciesRows = groupedBySpecies[speciesName];
            pointSpeciesList.push({
                type: 'data', label: speciesName, pointId: pointId, count: speciesRows.length,
                winnerSpecies: '', avgHeight: ''
            });
        });

        speciesSummary.push({ type: 'header', label: pointId, pointId, count: '', winnerSpecies: '', avgHeight: '' });
        speciesSummary.push(...pointSpeciesList);
    });

    const largeTrees = trees.filter(item => {
        const dbh = parseFloat(item.dbh);
        return !isNaN(dbh) && dbh >= 30;
    });
    const largeTreesByPoint = _.groupBy(largeTrees, 'pointId');
    const sortedLargeTrees = [];
    summaryPoints.forEach(pointId => {
        const treeList = largeTreesByPoint[pointId] || [];
        const pointData = groupedByPoint[pointId] || [];
        const sData = getClusterData(sMap, pointId);
        if (pointData.length === 0 && !sData.fclass && !sData.ftype) return;

        if (treeList.length > 0) {
            treeList.sort((a, b) => {
                const dbhA = parseFloat(a.dbh) || 0;
                const dbhB = parseFloat(b.dbh) || 0;
                if (dbhA !== dbhB) return dbhB - dbhA;
                return (a.species || '').localeCompare(b.species || '');
            }).forEach(item => {
                sortedLargeTrees.push({ 
                    pointId: item.pointId, species: item.species, dbh: item.dbh, 
                    combined: `${item.species}${item.dbh}`, dist: item.dist, azimuth: item.azimuth, note: item.note 
                });
            });
        } else {
            sortedLargeTrees.push({ pointId, species: '', dbh: '', combined: '', dist: '', azimuth: '', note: '' });
        }
    });

    return { speciesSummary, monitoringSummary, sortedLargeTrees };
};

                const readSheetData = (sheetKeywords, headerKeywords) => {
                    const actualSheetName = wb.SheetNames.find(name => {
                        const cleanName = clean(name);
                        return sheetKeywords.some(k => cleanName.includes(clean(k)));
                    });
                    if (!actualSheetName) return [];

                    const ws = wb.Sheets[actualSheetName];
                    updateRange(ws); // !!! 중요: 시트 범위 강제 재계산

                    const allRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
                    let headerIdx = -1;
                    
                    for (let i = 0; i < Math.min(allRows.length, 25); i++) {
                        if (!allRows[i]) continue;
                        const rowStr = allRows[i].map(c => clean(c)).join('|');
                        if (headerKeywords.some(k => rowStr.includes(clean(k)))) {
                            headerIdx = i;
                            break;
                        }
                    }
                    if (headerIdx === -1) return [];
                    return XLSX.utils.sheet_to_json(ws, { range: headerIdx, raw: false, defval: '' });
                };

                // 타입별 시트/헤더 키워드 설정
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
                let lastGenPid = '';
                rawGeneralJson.forEach(row => {
                    let pid = normalizeId(getCol(row, ['표본점번호', '표본점']));
                    if (!pid || pid === 'undefined') pid = lastGenPid;
                    else lastGenPid = pid;

                    const codeVal = getCol(row, ['토지이용정보', '토지이용']);
                    const code = codeVal !== undefined && codeVal !== null ? String(codeVal).trim() : '';
                    
                    if (pid && pid.length >= 5 && code && code !== 'undefined') {
                        generalMap[pid] = landUseCodes[code] || code;
                    }
                });

                // 3. 임분조사표 읽기
                const rawStandJson = readSheetData(standSheetKeywords, ['표본점', '임종', '임상']);
                const standMap = {};
                const forestClassCodes = { '0': '천연림', '1': '인공림' };
                const regenCodes = { '0': '기타', '1': '조림', '2': '천연하종', '3': '맹아' };
                const forestTypeCodes = { '0': '침엽수림', '1': '활엽수림', '2': '혼효림', '3': '비산림' };

                let lastStandPid = '';
                rawStandJson.forEach(row => {
                    let pid = normalizeId(getCol(row, ['표본점번호', '표본점']));
                    if (!pid || pid === 'undefined') pid = lastStandPid;
                    else lastStandPid = pid;

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

                // --- 2021자료.xlsx 연동 (전역 매핑 테이블 구축) ---
                fetch('/2021자료.xlsx')
                    .then(res => res.arrayBuffer())
                    .then(ab => {
                        const wb2021 = XLSX.read(ab, { type: 'array' });
                        
                        // 1) 2021DB요약 시트 읽기
                        const dbSheetName = wb2021.SheetNames.find(n => n.includes('2021DB요약')) || wb2021.SheetNames[0];
                        const dbData = XLSX.utils.sheet_to_json(wb2021.Sheets[dbSheetName]);
                        const dbLookup = {};
                        const idMap = {};
                        const clusterIdMap = {};

                        dbData.forEach(row => {
                            const newPid = String(row['표본점번호'] || '').trim();
                            const oldPid = String(row['구표본점번호'] || '').trim();
                            if (newPid) {
                                dbLookup[newPid] = row;
                                if (oldPid) {
                                    idMap[newPid] = oldPid;
                                    if (newPid.length >= 12 && oldPid.length >= 8) {
                                        clusterIdMap[newPid.slice(0, 12)] = oldPid.slice(0, 8);
                                    }
                                }
                            }
                        });

                        // 2) 임목조사표(2021) 시트 읽기 (트리 리스트 및 기타 정보용)
                        const treeSheetName = wb2021.SheetNames.find(n => n.includes('임목조사표(2021)')) || wb2021.SheetNames[1] || wb2021.SheetNames[0];
                        const treeData2021 = XLSX.utils.sheet_to_json(wb2021.Sheets[treeSheetName]);

                        // 지능형 ID 정규화 함수
                        const getNormalizedId = (pid) => {
                            const sPid = String(pid).trim();
                            if (idMap[sPid]) return idMap[sPid];
                            if (sPid.length === 13) {
                                const base = sPid.slice(0, 12);
                                if (clusterIdMap[base]) return clusterIdMap[base] + sPid.slice(-1);
                            }
                            return sPid;
                        };

                        // 2. 모든 데이터의 ID를 정규화
                        treeProcessed.forEach(t => { t.pointId = getNormalizedId(t.pointId); });
                        
                        const mappedGeneralMap = {};
                        Object.keys(generalMap).forEach(k => {
                            const normalizedKey = getNormalizedId(k);
                            if (!mappedGeneralMap[normalizedKey] || generalMap[k] === '경작지') {
                                mappedGeneralMap[normalizedKey] = generalMap[k];
                            }
                        });
                        
                        const mappedStandMap = {};
                        Object.keys(standMap).forEach(k => {
                            const normalizedKey = getNormalizedId(k);
                            mappedStandMap[normalizedKey] = standMap[k];
                        });

                        // 3. 결과 요약 생성
                        const res = generateSummaries(treeProcessed, mappedGeneralMap, mappedStandMap);
                        
                        res.monitoringSummary = res.monitoringSummary.map(row => {
                            if (row.landUse === '경작지' || row.ftype === '비산림') {
                                return { ...row, fclass: '', regen: '', ftype: '비산림', totalStems: 0, maxHSpecies: '', avgH: '' };
                            }
                            return row;
                        });

                        setData1(res.speciesSummary);
                        setDataMonitoring(res.monitoringSummary);
                        setData2(res.sortedLargeTrees);

                        // 4. 과거(2021) 데이터 요약
                        const currentPoints = new Set(res.monitoringSummary.map(r => r.pointId));
                        const filteredTree2021 = treeData2021.filter(item => {
                            const oldId = String(item['구표본점번호'] || '').trim();
                            const newId = String(item['표본점번호'] || '').trim();
                            return currentPoints.has(oldId) || currentPoints.has(newId);
                        });

                        const treeProcessed2021 = filteredTree2021.map(item => ({
                            pointId: String(item['구표본점번호'] || item['표본점번호'] || '').trim(),
                            species: String(item['수종명'] || '').trim(),
                            height: item['수고'],
                            dbh: item['흉고직경'],
                            dist: item['거리(m)'],
                            azimuth: item['방위각(º)'] || item['방위각(°)'],
                            note: String(item['비고(개체목구분코드)'] || '').replace('undefined', '').trim()
                        }));

                        const res2021 = generateSummaries(treeProcessed2021, {}, {}, Array.from(currentPoints).sort());
                        
                        // 2021DB요약의 값을 소계(subtotal) 행에 강제 주입
                        res2021.speciesSummary.forEach(row => {
                            if (row.type === 'subtotal') {
                                const dbInfo = dbLookup[row.pointId] || Object.values(dbLookup).find(val => val['구표본점번호'] === row.pointId);
                                if (dbInfo) {
                                    row.count = dbInfo['총본수'] !== undefined ? dbInfo['총본수'] : row.count;
                                    row.avgHeight = dbInfo['평균수고'] !== undefined ? dbInfo['평균수고'] : row.avgHeight;
                                }
                            }
                        });

                        setData2021_1(res2021.speciesSummary);
                        setData2021_2(res2021.sortedLargeTrees);

                        // 5. 모니터링비교_출현수종 생성 (전차기 수종 vs 현차기 수종 비교 분석)
                        const comparisonData = [];
                        const sortedPids = Array.from(currentPoints).sort();

                        sortedPids.forEach(pid => {
                            const prevSpecies = res2021.speciesSummary
                                .filter(s => s.pointId === pid && s.type === 'data')
                                .map(s => s.label)
                                .sort((a, b) => a.localeCompare(b));
                            const currSpecies = res.speciesSummary
                                .filter(s => s.pointId === pid && s.type === 'data')
                                .map(s => s.label)
                                .sort((a, b) => a.localeCompare(b));

                            const currUsed = new Set();
                            const plotRows = [];

                            // 1단계: 전차기(2021) 기준으로 매칭 시도
                            prevSpecies.forEach(pName => {
                                // 정확한 일치 확인
                                if (currSpecies.includes(pName)) {
                                    plotRows.push({ prev: pName, curr: pName });
                                    currUsed.add(pName);
                                } else {
                                    // 매칭되는 현차기 수종이 없음
                                    plotRows.push({ prev: pName, curr: '-' });
                                }
                            });

                            // 2단계: 현차기에만 새로 나타난 수종들을 하단에 추가 (가나다 순)
                            const newlyAdded = currSpecies.filter(c => !currUsed.has(c)).sort((a, b) => a.localeCompare(b));
                            newlyAdded.forEach(cName => {
                                plotRows.push({ prev: '', curr: cName });
                            });

                            // 3단계: 두 차기 모두 데이터가 없는 경우 (이미지의 276039603 케이스)
                            if (plotRows.length === 0) {
                                plotRows.push({ prev: '-', curr: '-' });
                            }

                            // 집락번호 및 전체 리스트에 추가
                            plotRows.forEach(item => {
                                comparisonData.push({
                                    clusterId: String(pid).length >= 5 ? String(pid).slice(0, -1) : pid,
                                    pointId: pid,
                                    prev: item.prev,
                                    curr: item.curr,
                                    cause: '',
                                    subCause: '',
                                    note: ''
                                });
                            });
                        });

                        setDataComparison(comparisonData);
                        setLoading(false);
                    })
                    .catch(err => {
                        console.error('2021 data error:', err);
                        const res = generateSummaries(treeProcessed, generalMap, standMap);
                        setData1(res.speciesSummary);
                        setDataMonitoring(res.monitoringSummary);
                        setData2(res.sortedLargeTrees);
                        setLoading(false);
                    });
            } catch (err) {
                console.error(err);
                setError(err.message || '파일 처리 중 오류가 발생했습니다.');
                setLoading(false);
            }
        };
        reader.readAsArrayBuffer(file);
    };

    const downloadExcel = async () => {
        const workbook = new ExcelJS.Workbook();
        
        // 헬퍼: 테두리 스타일
        const thinBorder = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        // 헬퍼: 열 너비 자동 조절
        const autoFit = (ws) => {
            ws.columns.forEach(column => {
                let maxColumnLength = 0;
                column.eachCell({ includeEmpty: true }, cell => {
                    const cellValue = cell.value ? cell.value.toString() : '';
                    let L = 0;
                    for (let i = 0; i < cellValue.length; i++) {
                        L += (cellValue.charCodeAt(i) > 127 ? 2.1 : 1.1);
                    }
                    if (L > maxColumnLength) maxColumnLength = L;
                });
                column.width = maxColumnLength < 10 ? 10 : maxColumnLength + 2;
            });
        };

        // 1. 모니터링 요약 시트
        const wsMon = workbook.addWorksheet('모니터링 요약');
        const monHeader = ['표본점번호', '토지이용', '임종', '갱신형태', '임상', '경급', '영급', '총본수', '평균 수고', '비산림면적(기본조사원)', '비산림면적(대경목조사원)'];
        const monHeaderRow = wsMon.addRow(monHeader);
        monHeaderRow.eachCell((cell, colNumber) => {
            cell.font = { bold: true };
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
            // 이미지처럼 홀수/짝수 열의 바탕색을 교차 적용 (홀수: 연한 녹색, 짝수: 흰색)
            const bgColor = colNumber % 2 === 1 ? 'FFE8F5E9' : 'FFFFFFFF';
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
            cell.border = thinBorder;
        });

        dataMonitoring.forEach(row => {
            const r = wsMon.addRow([
                row.pointId, row.landUse, row.fclass, row.regen, row.ftype, row.dclass, row.aclas,
                row.totalStems, row.avgH, row.nonForestBasic, row.nonForestLarge
            ]);
            r.eachCell((cell, colNumber) => { 
                cell.border = thinBorder; 
                cell.alignment = { vertical: 'middle' }; 
                // 헤더와 동일하게 홀수 열에만 바탕색 적용
                if (colNumber % 2 === 1) {
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };
                }
            });
        });
        autoFit(wsMon);

        // 2. 모니터링비교_출현수종 (핵심 스타일링 시트)
        if (dataComparison.length > 0) {
            const wsComp = workbook.addWorksheet('모니터링비교_출현수종');
            const compHeader = ['집락번호', '표본점번호', '전차기', '현차기', '변화원인', '세부변화원인', '비고'];
            const hRow = wsComp.addRow(compHeader);
            
            hRow.eachCell(cell => {
                cell.font = { bold: true };
                cell.alignment = { vertical: 'middle', horizontal: 'center' };
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2DCDB' } }; // 연한 핑크색
                cell.border = thinBorder;
            });

            let lastPid = null;
            let groupIdx = 0;
            
            dataComparison.forEach((item, idx) => {
                if (item.pointId !== lastPid) {
                    groupIdx++;
                    lastPid = item.pointId;
                }
                
                const row = wsComp.addRow([
                    item.clusterId, item.pointId, item.prev, item.curr, 
                    item.cause, item.subCause, item.note
                ]);
                
                row.eachCell((cell, colNumber) => {
                    cell.border = thinBorder;
                    cell.alignment = { vertical: 'middle', horizontal: 'center' };
                    
                    // 짝수 번째 표본점 그룹에 배경색 적용 (이미지처럼)
                    if (groupIdx % 2 === 0) {
                        cell.fill = {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: 'FFE8F5E9' } // 아주 연한 녹색
                        };
                    }
                });
            });

            // 셀 병합 로직 (집락번호 & 표본점번호)
            let clusterStart = 2;
            let pointStart = 2;
            for (let i = 0; i < dataComparison.length; i++) {
                const isLast = i === dataComparison.length - 1;
                const currentRow = i + 2;

                if (isLast || dataComparison[i].clusterId !== dataComparison[i + 1].clusterId) {
                    if (currentRow > clusterStart) wsComp.mergeCells(clusterStart, 1, currentRow, 1);
                    clusterStart = currentRow + 1;
                }
                if (isLast || dataComparison[i].pointId !== dataComparison[i + 1].pointId) {
                    if (currentRow > pointStart) wsComp.mergeCells(pointStart, 2, currentRow, 2);
                    pointStart = currentRow + 1;
                }
            }
            autoFit(wsComp);
        }

        // 3. 2021 출현종 요약
        if (data2021_1.length > 0) {
            const ws21 = workbook.addWorksheet('2021 출현종 요약');
            ws21.addRow(['표본점번호 및 수종명', '본수(개)', '', '']).font = { bold: true };
            data2021_1.forEach(row => {
                ws21.addRow([row.label, row.count, row.winnerSpecies, row.avgHeight]);
            });
            autoFit(ws21);
        }

        // 4. 2026 출현종 요약
        const ws26 = workbook.addWorksheet('2026 출현종 요약');
        ws26.addRow(['표본점번호 및 수종명', '본수(개)', '', '']).font = { bold: true };
        data1.forEach(row => {
            ws26.addRow([row.label, row.count, row.winnerSpecies, row.avgHeight]);
        });
        autoFit(ws26);

        // 5. 2021 대경목 출현 요약
        if (data2021_2.length > 0) {
            const ws21L = workbook.addWorksheet('2021 대경목 출현 요약');
            ws21L.addRow(['표본점번호', '수종명', '흉고직경', '수종명 흉고직경', '거리', '방위', '비고']).font = { bold: true };
            data2021_2.forEach(row => {
                ws21L.addRow([row.pointId, row.species, row.dbh, row.combined, row.dist, row.azimuth, row.note]);
            });
            autoFit(ws21L);
        }

        // 6. 2026 대경목 출현 요약
        const ws26L = workbook.addWorksheet('2026 대경목 출현 요약');
        ws26L.addRow(['표본점번호', '수종명', '흉고직경', '수종명 흉고직경', '거리', '방위', '비고']).font = { bold: true };
        data2.forEach(row => {
            ws26L.addRow([row.pointId, row.species, row.dbh, row.combined, row.dist, row.azimuth, row.note]);
        });
        autoFit(ws26L);

        // 파일 다운로드 실행
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = window.URL.createObjectURL(blob);
        const anchor = document.createElement('a');
        anchor.href = url;
        const originalBaseName = fileName.replace(/\.[^/.]+$/, "");
        anchor.download = `${originalBaseName}_모니터링 요약.xlsx`;
        anchor.click();
        window.URL.revokeObjectURL(url);
    };

    return (
        <div className="dashboard">
            <header>
                <h1>국가산림자원조사 결과보고서 자료 생성</h1>
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
                            className={`tab ${activeTab === 'comparison' ? 'active' : ''}`}
                            onClick={() => setActiveTab('comparison')}
                        >
                            모니터링비교_출현수종
                        </div>
                        <div
                            className={`tab ${activeTab === 'species2021' ? 'active' : ''}`}
                            onClick={() => setActiveTab('species2021')}
                        >
                            2021 출현종 요약
                        </div>
                        <div
                            className={`tab ${activeTab === 'species' ? 'active' : ''}`}
                            onClick={() => setActiveTab('species')}
                        >
                            2026 출현종 요약
                        </div>
                        <div
                            className={`tab ${activeTab === 'largeTrees2021' ? 'active' : ''}`}
                            onClick={() => setActiveTab('largeTrees2021')}
                        >
                            2021 대경목 출현 요약
                        </div>
                        <div
                            className={`tab ${activeTab === 'largeTrees' ? 'active' : ''}`}
                            onClick={() => setActiveTab('largeTrees')}
                        >
                            2026 대경목 출현 요약
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
                                    <th>평균 수고</th>
                                    <th>비산림면적(기본)</th>
                                    <th>비산림면적(대경목)</th>
                                </tr>
                            </thead>
                            <tbody>
                                {dataMonitoring.map((row, idx) => (
                                    <tr key={idx}>
                                        <td style={{ backgroundColor: '#E8F5E9' }}>{row.pointId}</td>
                                        <td>{row.landUse}</td>
                                        <td style={{ backgroundColor: '#E8F5E9' }}>{row.fclass}</td>
                                        <td>{row.regen}</td>
                                        <td style={{ backgroundColor: '#E8F5E9' }}>{row.ftype}</td>
                                        <td>{row.dclass}</td>
                                        <td style={{ backgroundColor: '#E8F5E9' }}>{row.aclas}</td>
                                        <td>{row.totalStems}</td>
                                        <td style={{ backgroundColor: '#E8F5E9' }}>{row.avgH}</td>
                                        <td>{row.nonForestBasic}</td>
                                        <td style={{ backgroundColor: '#E8F5E9' }}>{row.nonForestLarge}</td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>

                    <div className={`table-container ${activeTab !== 'species' ? 'hidden' : ''}`}>
                        {/* 메인 상세 테이블 */}
                        <table>
                            <thead>
                                <tr>
                                    <th>표본점번호 및 수종명</th>
                                    <th>본수(개)</th>
                                    <th></th>
                                    <th></th>
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

                    <div className={`table-container ${activeTab !== 'species2021' ? 'hidden' : ''}`}>
                        <table>
                            <thead>
                                <tr>
                                    <th>표본점번호 및 수종명</th>
                                    <th>본수(개)</th>
                                    <th></th>
                                    <th></th>
                                </tr>
                            </thead>
                            <tbody>
                                {data2021_1.map((row, idx) => (
                                    <tr key={idx} className={row.type === 'subtotal' ? 'subtotal' : row.type === 'header' ? 'point-header' : ''}>
                                        <td style={{ paddingLeft: row.type === 'data' ? '2.5rem' : '1rem' }}>
                                            {row.label}
                                        </td>
                                        <td>{row.count}</td>
                                        <td>{row.winnerSpecies}</td>
                                        <td>{row.avgHeight}</td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>

                    <div className={`table-container ${activeTab !== 'comparison' ? 'hidden' : ''}`}>
                        <table className="summary-table border-collapse w-full">
                            <thead style={{ backgroundColor: '#F2DCDB' }}>
                                <tr>
                                    <th>집락번호</th>
                                    <th>표본점번호</th>
                                    <th>전차기</th>
                                    <th>현차기</th>
                                    <th>변화원인</th>
                                    <th>세부변화원인</th>
                                    <th>비고</th>
                                </tr>
                            </thead>
                            <tbody>
                                {dataComparison.map((row, idx) => {
                                    const isFirstCluster = idx === 0 || dataComparison[idx - 1].clusterId !== row.clusterId;
                                    const isFirstPoint = idx === 0 || dataComparison[idx - 1].pointId !== row.pointId;
                                    return (
                                        <tr key={idx}>
                                            <td style={{ verticalAlign: 'top', fontWeight: isFirstCluster ? 'bold' : 'normal', color: isFirstCluster ? 'inherit' : 'transparent' }}>
                                                {isFirstCluster ? row.clusterId : ''}
                                            </td>
                                            <td style={{ verticalAlign: 'top', fontWeight: isFirstPoint ? 'bold' : 'normal', color: isFirstPoint ? 'inherit' : 'transparent' }}>
                                                {isFirstPoint ? row.pointId : ''}
                                            </td>
                                            <td>{row.prev}</td>
                                            <td>{row.curr}</td>
                                            <td>{row.cause}</td>
                                            <td>{row.subCause}</td>
                                            <td>{row.note}</td>
                                        </tr>
                                    );
                                })}
                            </tbody>
                        </table>
                    </div>

                    <div className={`table-container ${activeTab !== 'largeTrees2021' ? 'hidden' : ''}`}>
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
                                {data2021_2.map((row, idx) => (
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
