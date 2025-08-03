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

// ページ読み込み時にCSVファイルを読み込む
document.addEventListener('DOMContentLoaded', async function() {
    try {
        document.getElementById('loadingScreen').style.display = 'block';
        await loadCSVData();
        document.getElementById('loadingScreen').style.display = 'none';
        document.getElementById('startScreen').style.display = 'block';
    } catch (error) {
        showError('データの読み込みに失敗しました: ' + error.message);
    }
});

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
    document.getElementById('startScreen').style.display = 'none';
    document.getElementById('testScreen').style.display = 'block';
    
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
    questionStartTime = new Date();
    
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
    
    // 回答時間を記録
    const responseTime = new Date() - questionStartTime;
    
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
            responseTime: responseTime
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
    
    const totalTime = new Date() - testStartTime;
    const percentage = Math.round((score / testData.length) * 100);
    
    document.getElementById('finalScore').textContent = `${score} / ${testData.length}`;
    document.getElementById('percentage').textContent = percentage;
    
    // 結果をコンソールに出力（デバッグ用）
    console.log('テスト結果:', {
        score: score,
        totalQuestions: testData.length,
        percentage: percentage,
        totalTime: Math.round(totalTime / 1000) + '秒',
        results: testResults
    });
    
    // 結果をローカルストレージに保存
    saveResults();
}

// 結果の保存
function saveResults() {
    const resultData = {
        date: new Date().toISOString(),
        score: score,
        totalQuestions: testData.length,
        percentage: Math.round((score / testData.length) * 100),
        details: testResults
    };
    
    // 既存の結果を取得
    let allResults = JSON.parse(localStorage.getItem('listeningTestResults') || '[]');
    allResults.push(resultData);
    
    // 最新の10件のみ保存
    if (allResults.length > 10) {
        allResults = allResults.slice(-10);
    }
    
    localStorage.setItem('listeningTestResults', JSON.stringify(allResults));
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
