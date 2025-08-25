// グローバル変数
let testData = [];
let practiceData = [];
let currentQuestion = 0;
let score = 0;
let practiceScore = 0;
let audioPlayed = false;
let selectedAnswer = null;
let currentAudio = null;
let testStartTime = null;
let questionStartTime = null;
let testResults = [];
let isPractice = true;
let currentData = [];
let participantName = ''; // 受験者名を保存

// ページ読み込み時にCSVファイルを読み込む
document.addEventListener('DOMContentLoaded', async function() {
    try {
        document.getElementById('loadingScreen').style.display = 'block';
        await loadCSVData();
        document.getElementById('loadingScreen').style.display = 'none';
        document.getElementById('startScreen').style.display = 'block';
        
        // 名前入力欄のイベントリスナーを追加
        setupNameInputListeners();
    } catch (error) {
        showError('データの読み込みに失敗しました: ' + error.message);
    }
});

// 名前入力欄のイベントリスナー設定
function setupNameInputListeners() {
    const nameInput = document.getElementById('participantName');
    if (nameInput) {
        // Enterキーで開始
        nameInput.addEventListener('keypress', function(event) {
            if (event.key === 'Enter') {
                startPractice();
            }
        });
        
        // 入力時にエラーをクリア
        nameInput.addEventListener('input', function() {
            if (this.value.trim() !== '') {
                document.getElementById('nameError').style.display = 'none';
                this.style.borderColor = '#4CAF50';
            }
        });
    }
}

// 名前入力の検証
function validateName() {
    const nameInput = document.getElementById('participantName');
    const nameError = document.getElementById('nameError');
    const name = nameInput.value.trim();
    
    if (name === '') {
        nameError.style.display = 'block';
        nameInput.style.borderColor = '#d32f2f';
        return false;
    } else {
        nameError.style.display = 'none';
        nameInput.style.borderColor = '#4CAF50';
        participantName = name;
        return true;
    }
}

// CSVファイルの読み込み
async function loadCSVData() {
    try {
        const response = await fetch('stimuli_list.csv');
        if (!response.ok) {
            throw new Error('CSVファイルが見つかりません');
        }
        
        // UTF-8でテキストを読み込む
        const text = await response.text();
        
        // CSVをパース
        const rows = parseCSV(text);
        
        // ヘッダー行を取得
        const headers = rows[0];
        
        // データを練習問題と本番問題に分ける
        practiceData = [];
        testData = [];
        
        // ヘッダー行をスキップして処理
        for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            if (row.length < headers.length) continue;
            
            // 各列のインデックスを取得
            const displayIndex = headers.indexOf('display');
            const answerIndex = headers.indexOf('ANSWER');
            const stimuliIndex = headers.indexOf('stimuli');
            const responseAIndex = headers.indexOf('responseA');
            const responseBIndex = headers.indexOf('responseB');
            const responseCIndex = headers.indexOf('responseC');
            const responseDIndex = headers.indexOf('responseD');
            const targetwordIndex = headers.indexOf('targetword');
            
            const display = row[displayIndex];
            const stimuli = row[stimuliIndex];
            
            if (!display || !stimuli || display === '') continue;
            
            const choices = [
                row[responseAIndex],
                row[responseBIndex],
                row[responseCIndex],
                row[responseDIndex]
            ];
            
            const answer = row[answerIndex];
            
            // 正解のインデックスを見つける
            let correctIndex = -1;
            for (let j = 0; j < choices.length; j++) {
                if (choices[j] === answer) {
                    correctIndex = j;
                    break;
                }
            }
            
            const questionData = {
                audio: display.startsWith('Practice') 
                    ? `practice_stimuli/${stimuli}`
                    : `stimuli/${stimuli}`,
                choices: choices,
                correct: correctIndex,
                targetWord: row[targetwordIndex]
            };
            
            if (display.startsWith('Practice')) {
                practiceData.push(questionData);
            } else if (display === 'Trial') {
                testData.push(questionData);
            }
        }
        
        console.log(`データ読み込み完了: 練習問題 ${practiceData.length}問, 本番問題 ${testData.length}問`);
        
        if (practiceData.length === 0 || testData.length === 0) {
            throw new Error('有効な問題データが見つかりません');
        }
        
    } catch (error) {
        console.error('CSVファイル読み込みエラー:', error);
        throw error;
    }
}

// CSVパーサー（日本語対応）
function parseCSV(text) {
    const rows = [];
    const lines = text.split(/\r?\n/);
    
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        if (line === '') continue;
        
        const row = [];
        let cell = '';
        let inQuotes = false;
        
        for (let j = 0; j < line.length; j++) {
            const char = line[j];
            const nextChar = line[j + 1];
            
            if (char === '"') {
                if (inQuotes && nextChar === '"') {
                    cell += '"';
                    j++; // Skip next quote
                } else {
                    inQuotes = !inQuotes;
                }
            } else if (char === ',' && !inQuotes) {
                row.push(cell);
                cell = '';
            } else {
                cell += char;
            }
        }
        
        row.push(cell); // Add last cell
        rows.push(row);
    }
    
    return rows;
}

// エラー表示
function showError(message) {
    document.getElementById('loadingScreen').style.display = 'none';
    document.getElementById('startScreen').style.display = 'none';
    document.getElementById('errorMessage').style.display = 'block';
    document.getElementById('errorText').textContent = message;
}

// 練習開始
function startPractice() {
    // 名前の検証
    if (!validateName()) {
        return;
    }
    
    document.getElementById('startScreen').style.display = 'none';
    document.getElementById('testScreen').style.display = 'block';
    
    // 名前を表示
    if (document.getElementById('testParticipantName')) {
        document.getElementById('testParticipantName').textContent = participantName;
    }
    
    isPractice = true;
    currentData = practiceData;
    currentQuestion = 0;
    practiceScore = 0;
    
    document.getElementById('questionType').textContent = '練習問題';
    loadQuestion();
}

// 本番テスト開始
function startMainTest() {
    document.getElementById('practiceCompleteScreen').style.display = 'none';
    document.getElementById('testScreen').style.display = 'block';
    
    // 名前を表示
    if (document.getElementById('testParticipantName')) {
        document.getElementById('testParticipantName').textContent = participantName;
    }
    
    isPractice = false;
    currentData = testData;
    currentQuestion = 0;
    score = 0;
    testResults = [];
    testStartTime = new Date();
    
    document.getElementById('questionType').textContent = '本番テスト';
    loadQuestion();
}

// 問題の読み込み
function loadQuestion() {
    const question = currentData[currentQuestion];
    questionStartTime = performance.now(); // より正確なタイミング測定
    
    // 進捗を更新
    const totalQuestions = currentData.length;
    document.getElementById('progress').textContent = 
        `問題 ${currentQuestion + 1} / ${totalQuestions}`;
    
    // 音声の準備
    if (currentAudio) {
        currentAudio.pause();
        currentAudio = null;
    }
    
    currentAudio = new Audio(question.audio);
    
    // 音声読み込みエラーの処理
    currentAudio.addEventListener('error', function(e) {
        console.error('音声ファイルの読み込みエラー:', e);
        document.getElementById('audioStatus').textContent = 
            '音声ファイルの読み込みに失敗しました。';
        document.getElementById('playButton').disabled = true;
    });
    
    audioPlayed = false;
    selectedAnswer = null;
    
    // ボタンをリセット
    document.getElementById('playButton').disabled = false;
    document.getElementById('playButton').textContent = '音声を再生（1回のみ）';
    document.getElementById('audioStatus').textContent = '';
    document.getElementById('nextButton').style.display = 'none';
    
    // 選択肢をシャッフルして表示
    const shuffledIndices = shuffleArray([0, 1, 2, 3]);
    const choiceButtons = document.querySelectorAll('.choice-button');
    
    shuffledIndices.forEach((originalIndex, buttonIndex) => {
        choiceButtons[buttonIndex].textContent = question.choices[originalIndex];
        choiceButtons[buttonIndex].setAttribute('data-original-index', originalIndex);
        choiceButtons[buttonIndex].className = 'choice-button';
        choiceButtons[buttonIndex].disabled = true;
    });
}

// 配列のシャッフル（Fisher-Yates アルゴリズム）
function shuffleArray(array) {
    const shuffled = [...array];
    for (let i = shuffled.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [shuffled[i], shuffled[j]] = [shuffled[j], shuffled[i]];
    }
    return shuffled;
}

// 音声再生
function playAudio() {
    if (!audioPlayed && currentAudio) {
        currentAudio.play()
            .then(() => {
                audioPlayed = true;
                document.getElementById('playButton').disabled = true;
                document.getElementById('playButton').textContent = '再生済み';
                document.getElementById('audioStatus').textContent = 
                    '音声を再生しました。選択肢を選んでください。';
                
                // 選択肢を有効化
                const choiceButtons = document.querySelectorAll('.choice-button');
                choiceButtons.forEach(button => {
                    button.disabled = false;
                });
            })
            .catch(error => {
                console.error('音声再生エラー:', error);
                document.getElementById('audioStatus').textContent = 
                    '音声の再生に失敗しました。ブラウザの設定を確認してください。';
            });
    }
}

// 選択肢の選択
function selectChoice(buttonIndex) {
    if (!audioPlayed || selectedAnswer !== null) return;
    
    selectedAnswer = buttonIndex;
    const choiceButtons = document.querySelectorAll('.choice-button');
    const selectedButton = choiceButtons[buttonIndex];
    const originalIndex = parseInt(selectedButton.getAttribute('data-original-index'));
    const question = currentData[currentQuestion];
    
    // 回答時間を記録（ミリ秒単位）
    const responseTime = Math.round(performance.now() - questionStartTime);
    
    // 選択したボタンをマーク
    selectedButton.classList.add('selected');
    
    // 正解判定
    let isCorrect = false;
    if (originalIndex === question.correct) {
        selectedButton.classList.add('correct');
        if (isPractice) {
            practiceScore++;
        } else {
            score++;
        }
        isCorrect = true;
    } else {
        selectedButton.classList.add('incorrect');
        // 正解を表示
        choiceButtons.forEach(button => {
            if (parseInt(button.getAttribute('data-original-index')) === question.correct) {
                button.classList.add('correct');
            }
        });
    }
    
    // 本番テストの場合のみ結果を記録
    if (!isPractice) {
        testResults.push({
            questionNumber: currentQuestion + 1,
            targetWord: question.targetWord,
            selectedAnswer: question.choices[originalIndex],
            correctAnswer: question.choices[question.correct],
            isCorrect: isCorrect,
            responseTime: responseTime  // ミリ秒で保存
        });
    }
    
    // 選択肢を無効化
    choiceButtons.forEach(button => {
        button.disabled = true;
    });
    
    // 次へボタンを表示
    document.getElementById('nextButton').style.display = 'block';
}

// 次の問題へ
function nextQuestion() {
    currentQuestion++;
    
    if (currentQuestion < currentData.length) {
        loadQuestion();
    } else {
        if (isPractice) {
            // 練習終了
            showPracticeComplete();
        } else {
            // 本番終了
            showResults();
        }
    }
}

// 練習完了画面を表示
function showPracticeComplete() {
    document.getElementById('testScreen').style.display = 'none';
    document.getElementById('practiceCompleteScreen').style.display = 'block';
    document.getElementById('practiceScore').textContent = practiceScore;
    
    // 名前を表示
    if (document.getElementById('practiceParticipantName')) {
        document.getElementById('practiceParticipantName').textContent = participantName;
    }
    
    // 音声を停止
    if (currentAudio) {
        currentAudio.pause();
        currentAudio = null;
    }
}

// 結果表示
function showResults() {
    document.getElementById('testScreen').style.display = 'none';
    document.getElementById('resultScreen').style.display = 'block';
    
    const totalTime = performance.now() - testStartTime;
    const percentage = Math.round((score / testData.length) * 100);
    
    // 名前を表示
    if (document.getElementById('resultParticipantName')) {
        document.getElementById('resultParticipantName').textContent = participantName;
    }
    
    // テスト日時を表示
    const now = new Date();
    const dateTimeString = now.toLocaleString('ja-JP', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit'
    });
    if (document.getElementById('testDateTime')) {
        document.getElementById('testDateTime').textContent = dateTimeString;
    }
    
    document.getElementById('finalScore').textContent = `${score} / ${testData.length}`;
    document.getElementById('percentage').textContent = percentage;
    
    // 履歴統計を表示
    displayHistoryStats();
    
    // 結果をコンソールに出力（デバッグ用）
    console.log('テスト結果:', {
        participantName: participantName,
        score: score,
        totalQuestions: testData.length,
        percentage: percentage,
        totalTime: Math.round(totalTime) + 'ミリ秒',
        results: testResults
    });
    
    // 結果をローカルストレージに保存
    saveResults();
}

// 履歴統計の表示
function displayHistoryStats() {
    const allResults = JSON.parse(localStorage.getItem('listeningTestResults') || '[]');
    
    // 現在の受験者の結果のみをフィルタリング
    const userResults = allResults.filter(r => r.participantName === participantName);
    
    // 受験回数を表示（今回を含む）
    const totalAttempts = userResults.length + 1;
    if (document.getElementById('totalAttempts')) {
        document.getElementById('totalAttempts').textContent = totalAttempts;
    }
    
    // 平均正答率を計算（今回の結果を含む）
    if (userResults.length > 0) {
        const currentPercentage = Math.round((score / testData.length) * 100);
        const totalPercentage = userResults.reduce((sum, r) => sum + r.percentage, 0) + currentPercentage;
        const averagePercentage = Math.round(totalPercentage / totalAttempts);
        if (document.getElementById('averageScore')) {
            document.getElementById('averageScore').textContent = averagePercentage;
        }
    } else {
        // 初回の場合
        const currentPercentage = Math.round((score / testData.length) * 100);
        if (document.getElementById('averageScore')) {
            document.getElementById('averageScore').textContent = currentPercentage;
        }
    }
}

// 結果の保存
function saveResults() {
    const now = new Date();
    const resultData = {
        participantName: participantName,
        date: now.toISOString(),
        dateString: now.toLocaleString('ja-JP'),
        score: score,
        totalQuestions: testData.length,
        percentage: Math.round((score / testData.length) * 100),
        details: testResults
    };
    
    // 既存の結果を取得
    let allResults = JSON.parse(localStorage.getItem('listeningTestResults') || '[]');
    allResults.push(resultData);
    
    // 最新の100件のみ保存
    if (allResults.length > 100) {
        allResults = allResults.slice(-100);
    }
    
    localStorage.setItem('listeningTestResults', JSON.stringify(allResults));
    
    console.log(`保存済みの結果: ${allResults.length}件`);
    console.log('最新の結果:', resultData);
}

// 全受験履歴をExcelでダウンロード（複数シート）
function downloadAllResults() {
    const allResults = JSON.parse(localStorage.getItem('listeningTestResults') || '[]');
    
    // 現在のテスト結果も含める
    const currentResult = {
        participantName: participantName,
        dateString: new Date().toLocaleString('ja-JP'),
        score: score,
        totalQuestions: testData.length,
        percentage: Math.round((score / testData.length) * 100),
        details: testResults
    };
    
    // Excelワークブックを作成
    const wb = XLSX.utils.book_new();
    
    // シート1: サマリー
    const summaryData = [];
    
    // タイトルと基本情報
    summaryData.push(['英単語リスニングテスト 全受験履歴']);
    summaryData.push(['ダウンロード日時', new Date().toLocaleString('ja-JP')]);
    summaryData.push([]);
    
    // 全体サマリー
    summaryData.push(['【全体サマリー】']);
    const totalTests = allResults.length + 1;
    const allScores = [...allResults.map(r => r.percentage), currentResult.percentage];
    const avgScore = Math.round(allScores.reduce((sum, s) => sum + s, 0) / allScores.length);
    const maxScore = Math.max(...allScores);
    const minScore = Math.min(...allScores);
    
    summaryData.push(['総受験回数', totalTests]);
    summaryData.push(['全体平均正答率', avgScore + '%']);
    summaryData.push(['最高正答率', maxScore + '%']);
    summaryData.push(['最低正答率', minScore + '%']);
    summaryData.push([]);
    
    // 受験者別サマリー
    summaryData.push(['【受験者別サマリー】']);
    summaryData.push(['受験者名', '受験回数', '平均正答率(%)', '最高正答率(%)', '最低正答率(%)']);
    
    const participantStats = {};
    [...allResults, currentResult].forEach(result => {
        const name = result.participantName || '名前未入力';
        if (!participantStats[name]) {
            participantStats[name] = {
                count: 0,
                scores: []
            };
        }
        participantStats[name].count++;
        participantStats[name].scores.push(result.percentage);
    });
    
    Object.keys(participantStats).forEach(name => {
        const stats = participantStats[name];
        const avgScore = Math.round(stats.scores.reduce((sum, s) => sum + s, 0) / stats.scores.length);
        const maxScore = Math.max(...stats.scores);
        const minScore = Math.min(...stats.scores);
        summaryData.push([name, stats.count, avgScore, maxScore, minScore]);
    });
    
    const ws_summary = XLSX.utils.aoa_to_sheet(summaryData);
    
    // 列幅の設定
    ws_summary['!cols'] = [
        { wch: 25 },  // 受験者名
        { wch: 15 },  // 受験回数
        { wch: 18 },  // 平均正答率
        { wch: 18 },  // 最高正答率
        { wch: 18 }   // 最低正答率
    ];
    
    XLSX.utils.book_append_sheet(wb, ws_summary, 'サマリー');
    
    // シート2: 全受験履歴
    const historyData = [];
    historyData.push(['番号', '受験者名', '日時', '正答数', '総問題数', '正答率(%)']);
    
    [...allResults, currentResult].forEach((result, index) => {
        historyData.push([
            index + 1,
            result.participantName || '名前未入力',
            result.dateString,
            result.score,
            result.totalQuestions,
            result.percentage
        ]);
    });
    
    const ws_history = XLSX.utils.aoa_to_sheet(historyData);
    
    // 列幅の設定
    ws_history['!cols'] = [
        { wch: 8 },   // 番号
        { wch: 20 },  // 受験者名
        { wch: 25 },  // 日時
        { wch: 10 },  // 正答数
        { wch: 12 },  // 総問題数
        { wch: 12 }   // 正答率
    ];
    
    XLSX.utils.book_append_sheet(wb, ws_history, '全受験履歴');
    
    // シート3: 最新テストの詳細（80問）
    const detailData = [];
    detailData.push(['受験者', currentResult.participantName]);
    detailData.push(['日時', currentResult.dateString]);
    detailData.push(['結果', `${currentResult.score} / ${currentResult.totalQuestions} (${currentResult.percentage}%)`]);
    detailData.push([]);
    detailData.push(['問題番号', '英単語', '選択した答え', '正解', '正誤', '回答時間(ミリ秒)', '回答時間(秒)']);
    
    if (currentResult.details && currentResult.details.length > 0) {
        currentResult.details.forEach(detail => {
            detailData.push([
                detail.questionNumber,
                detail.targetWord || '',
                detail.selectedAnswer,
                detail.correctAnswer,
                detail.isCorrect ? '○' : '×',
                detail.responseTime,  // ミリ秒単位
                (detail.responseTime / 1000).toFixed(2)  // 秒単位（小数点2桁）
            ]);
        });
    }
    
    const ws_detail = XLSX.utils.aoa_to_sheet(detailData);
    
    // 列幅の設定
    ws_detail['!cols'] = [
        { wch: 10 },  // 問題番号
        { wch: 15 },  // 英単語
        { wch: 20 },  // 選択した答え
        { wch: 20 },  // 正解
        { wch: 8 },   // 正誤
        { wch: 18 },  // 回答時間(ミリ秒)
        { wch: 15 }   // 回答時間(秒)
    ];
    
    XLSX.utils.book_append_sheet(wb, ws_detail, '最新テスト詳細');
    
    // シート4: 統計分析
    const statsData = [];
    statsData.push(['【最新テストの統計分析】']);
    statsData.push([]);
    
    if (currentResult.details && currentResult.details.length > 0) {
        // 基本統計
        const avgResponseTime = currentResult.details.reduce((sum, d) => sum + d.responseTime, 0) / currentResult.details.length;
        statsData.push(['平均回答時間（ミリ秒）', Math.round(avgResponseTime)]);
        statsData.push(['平均回答時間（秒）', (avgResponseTime / 1000).toFixed(2)]);
        statsData.push([]);
        
        // 正答・誤答別統計
        const correctAnswers = currentResult.details.filter(d => d.isCorrect);
        const incorrectAnswers = currentResult.details.filter(d => !d.isCorrect);
        
        statsData.push(['正答問題数', correctAnswers.length]);
        statsData.push(['誤答問題数', incorrectAnswers.length]);
        statsData.push([]);
        
        if (correctAnswers.length > 0) {
            const avgCorrectTime = correctAnswers.reduce((sum, d) => sum + d.responseTime, 0) / correctAnswers.length;
            statsData.push(['正答時の平均回答時間（ミリ秒）', Math.round(avgCorrectTime)]);
            statsData.push(['正答時の平均回答時間（秒）', (avgCorrectTime / 1000).toFixed(2)]);
            
            // 最速・最遅
            const fastestCorrect = Math.min(...correctAnswers.map(d => d.responseTime));
            const slowestCorrect = Math.max(...correctAnswers.map(d => d.responseTime));
            statsData.push(['正答時の最速回答（ミリ秒）', fastestCorrect]);
            statsData.push(['正答時の最遅回答（ミリ秒）', slowestCorrect]);
            
            // 中央値
            const sortedCorrectTimes = correctAnswers.map(d => d.responseTime).sort((a, b) => a - b);
            const medianCorrect = sortedCorrectTimes.length % 2 === 0
                ? (sortedCorrectTimes[sortedCorrectTimes.length / 2 - 1] + sortedCorrectTimes[sortedCorrectTimes.length / 2]) / 2
                : sortedCorrectTimes[Math.floor(sortedCorrectTimes.length / 2)];
            statsData.push(['正答時の中央値（ミリ秒）', Math.round(medianCorrect)]);
        }
        
        statsData.push([]);
        
        if (incorrectAnswers.length > 0) {
            const avgIncorrectTime = incorrectAnswers.reduce((sum, d) => sum + d.responseTime, 0) / incorrectAnswers.length;
            statsData.push(['誤答時の平均回答時間（ミリ秒）', Math.round(avgIncorrectTime)]);
            statsData.push(['誤答時の平均回答時間（秒）', (avgIncorrectTime / 1000).toFixed(2)]);
            
            // 最速・最遅
            const fastestIncorrect = Math.min(...incorrectAnswers.map(d => d.responseTime));
            const slowestIncorrect = Math.max(...incorrectAnswers.map(d => d.responseTime));
            statsData.push(['誤答時の最速回答（ミリ秒）', fastestIncorrect]);
            statsData.push(['誤答時の最遅回答（ミリ秒）', slowestIncorrect]);
            
            // 中央値
            const sortedIncorrectTimes = incorrectAnswers.map(d => d.responseTime).sort((a, b) => a - b);
            const medianIncorrect = sortedIncorrectTimes.length % 2 === 0
                ? (sortedIncorrectTimes[sortedIncorrectTimes.length / 2 - 1] + sortedIncorrectTimes[sortedIncorrectTimes.length / 2]) / 2
                : sortedIncorrectTimes[Math.floor(sortedIncorrectTimes.length / 2)];
            statsData.push(['誤答時の中央値（ミリ秒）', Math.round(medianIncorrect)]);
        }
        
        statsData.push([]);
        statsData.push(['【回答時間分布】']);
        statsData.push(['回答時間範囲', '問題数', '割合(%)', '正答数', '誤答数']);
        
        // 回答時間の分布（より細かく）
        const timeRanges = [
            { label: '0-1秒', min: 0, max: 1000, count: 0, correct: 0, incorrect: 0 },
            { label: '1-2秒', min: 1000, max: 2000, count: 0, correct: 0, incorrect: 0 },
            { label: '2-3秒', min: 2000, max: 3000, count: 0, correct: 0, incorrect: 0 },
            { label: '3-4秒', min: 3000, max: 4000, count: 0, correct: 0, incorrect: 0 },
            { label: '4-5秒', min: 4000, max: 5000, count: 0, correct: 0, incorrect: 0 },
            { label: '5-6秒', min: 5000, max: 6000, count: 0, correct: 0, incorrect: 0 },
            { label: '6-7秒', min: 6000, max: 7000, count: 0, correct: 0, incorrect: 0 },
            { label: '7-8秒', min: 7000, max: 8000, count: 0, correct: 0, incorrect: 0 },
            { label: '8-9秒', min: 8000, max: 9000, count: 0, correct: 0, incorrect: 0 },
            { label: '9-10秒', min: 9000, max: 10000, count: 0, correct: 0, incorrect: 0 },
            { label: '10秒以上', min: 10000, max: Infinity, count: 0, correct: 0, incorrect: 0 }
        ];
        
        currentResult.details.forEach(detail => {
            const time = detail.responseTime;
            const range = timeRanges.find(r => time >= r.min && time < r.max);
            if (range) {
                range.count++;
                if (detail.isCorrect) {
                    range.correct++;
                } else {
                    range.incorrect++;
                }
            }
        });
        
        timeRanges.forEach(range => {
            if (range.count > 0) {
                const percentage = Math.round((range.count / currentResult.details.length) * 100);
                statsData.push([range.label, range.count, percentage, range.correct, range.incorrect]);
            }
        });
        
        statsData.push([]);
        statsData.push(['【標準偏差】']);
        
        // 標準偏差の計算
        const mean = avgResponseTime;
        const variance = currentResult.details.reduce((sum, d) => sum + Math.pow(d.responseTime - mean, 2), 0) / currentResult.details.length;
        const stdDev = Math.sqrt(variance);
        
        statsData.push(['全体の標準偏差（ミリ秒）', Math.round(stdDev)]);
        statsData.push(['全体の標準偏差（秒）', (stdDev / 1000).toFixed(2)]);
    }
    
    const ws_stats = XLSX.utils.aoa_to_sheet(statsData);
    
    // 列幅の設定
    ws_stats['!cols'] = [
        { wch: 35 },  // 項目名
        { wch: 15 },  // 値
        { wch: 12 },  // 割合
        { wch: 10 },  // 正答数
        { wch: 10 }   // 誤答数
    ];
    
    XLSX.utils.book_append_sheet(wb, ws_stats, '統計分析');
    
    // Excelファイルとして保存
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);
    const filename = `listening_test_results_${timestamp}.xlsx`;
    
    XLSX.writeFile(wb, filename);
    
    console.log(`Excelファイル "${filename}" をダウンロードしました`);
}

// テストの再開（最初から）
function restartTest() {
    currentQuestion = 0;
    score = 0;
    practiceScore = 0;
    testResults = [];
    selectedAnswer = null;
    audioPlayed = false;
    isPractice = true;
    participantName = '';  // 名前もリセット
    
    // 名前入力欄をクリア
    const nameInput = document.getElementById('participantName');
    if (nameInput) {
        nameInput.value = '';
        nameInput.style.borderColor = '#ddd';
    }
    
    const nameError = document.getElementById('nameError');
    if (nameError) {
        nameError.style.display = 'none';
    }
    
    if (currentAudio) {
        currentAudio.pause();
        currentAudio = null;
    }
    
    document.getElementById('resultScreen').style.display = 'none';
    document.getElementById('startScreen').style.display = 'block';
}

// キーボードショートカット
document.addEventListener('keydown', function(event) {
    // スペースキーで音声再生
    if (event.code === 'Space' && !audioPlayed && 
        document.getElementById('testScreen').style.display === 'block') {
        event.preventDefault();
        playAudio();
    }
    
    // 数字キー1-4で選択肢を選択
    if (audioPlayed && selectedAnswer === null && 
        document.getElementById('testScreen').style.display === 'block') {
        const key = parseInt(event.key);
        if (key >= 1 && key <= 4) {
            selectChoice(key - 1);
        }
    }
    
    // Enterキーで次の問題へ
    if (event.code === 'Enter' && selectedAnswer !== null && 
        document.getElementById('nextButton').style.display === 'block') {
        nextQuestion();
    }
});

// 結果履歴の表示（デバッグ用）
function showResultHistory() {
    const allResults = JSON.parse(localStorage.getItem('listeningTestResults') || '[]');
    console.log('=== 保存されている全結果 ===');
    allResults.forEach((result, index) => {
        console.log(`${index + 1}. ${result.participantName} - ${result.dateString} - ${result.score}/${result.totalQuestions} (${result.percentage}%)`);
    });
}

// データクリア関数（デバッグ用）
function clearAllData() {
    if (confirm('本当にすべての履歴データを削除しますか？')) {
        localStorage.removeItem('listeningTestResults');
        console.log('すべての履歴データを削除しました');
    }
}
