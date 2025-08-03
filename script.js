// テストデータ（実際のデータに置き換えてください）
const testData = [
    {
        audio: "stimuli/101S1F01.wav",
        choices: ["感謝する", "主張する", "卒業する", "勉強する"],
        correct: 0,
        targetWord: "appreciate"
    },
    {
        audio: "stimuli/102S1F01.wav",
        choices: ["幹部", "建築", "執行", "学校"],
        correct: 0,
        targetWord: "executive"
    },
    {
        audio: "stimuli/103S1F01.wav",
        choices: ["歓迎会", "保険", "旅行", "製造業"],
        correct: 0,
        targetWord: "reception"
    },
    {
        audio: "stimuli/104S1F01.wav",
        choices: ["契約", "遠足", "軽食", "証明書"],
        correct: 2,
        targetWord: "refreshment"
    },
    {
        audio: "stimuli/105S1F01.wav",
        choices: ["比較する", "出現する", "要求する", "評価する"],
        correct: 2,
        targetWord: "require"
    }
    // ... 残りの75問のデータを追加
    // 実際の実装では、Excelファイルのデータを変換して使用してください
];

// グローバル変数
let currentQuestion = 0;
let score = 0;
let audioPlayed = false;
let selectedAnswer = null;
let currentAudio = null;
let testStartTime = null;
let questionStartTime = null;
let testResults = [];

// 練習問題用のデータ（必要に応じて追加）
const practiceData = [
    {
        audio: "practice_stimuli/Practice01F01.wav",
        choices: ["勉強する", "補足する", "配管工", "ネコ"],
        correct: 0,
        targetWord: "study"
    },
    {
        audio: "practice_stimuli/Practice02F01.wav",
        choices: ["コーヒー", "邪悪", "授業料", "警告"],
        correct: 0,
        targetWord: "coffee"
    },
    {
        audio: "practice_stimuli/Practice03F01.wav",
        choices: ["陶器", "部長", "配管工", "ネコ"],
        correct: 3,
        targetWord: "cat"
    }
];

// テスト開始
function startTest() {
    document.getElementById('startScreen').style.display = 'none';
    document.getElementById('testScreen').style.display = 'block';
    testStartTime = new Date();
    loadQuestion();
}

// 問題の読み込み
function loadQuestion() {
    const question = testData[currentQuestion];
    questionStartTime = new Date();
    
    // 進捗を更新
    document.getElementById('progress').textContent = 
        `問題 ${currentQuestion + 1} / ${testData.length}`;
    
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
    const question = testData[currentQuestion];
    
    // 回答時間を記録
    const responseTime = new Date() - questionStartTime;
    
    // 選択したボタンをマーク
    selectedButton.classList.add('selected');
    
    // 正解判定
    let isCorrect = false;
    if (originalIndex === question.correct) {
        selectedButton.classList.add('correct');
        score++;
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
    
    // 結果を記録
    testResults.push({
        questionNumber: currentQuestion + 1,
        targetWord: question.targetWord,
        selectedAnswer: question.choices[originalIndex],
        correctAnswer: question.choices[question.correct],
        isCorrect: isCorrect,
        responseTime: responseTime
    });
    
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
    
    if (currentQuestion < testData.length) {
        loadQuestion();
    } else {
        showResults();
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
    
    // 結果をローカルストレージに保存（オプション）
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

// テストの再開
function restartTest() {
    currentQuestion = 0;
    score = 0;
    testResults = [];
    selectedAnswer = null;
    audioPlayed = false;
    
    if (currentAudio) {
        currentAudio.pause();
        currentAudio = null;
    }
    
    document.getElementById('resultScreen').style.display = 'none';
    document.getElementById('startScreen').style.display = 'block';
}

// ページ読み込み時の初期化
document.addEventListener('DOMContentLoaded', function() {
    // 音声ファイルのプリロード（オプション）
    preloadAudioFiles();
});

// 音声ファイルのプリロード
function preloadAudioFiles() {
    console.log('音声ファイルのプリロードを開始します...');
    
    // 最初の数問だけプリロード
    const preloadCount = Math.min(5, testData.length);
    for (let i = 0; i < preloadCount; i++) {
        const audio = new Audio(testData[i].audio);
        audio.preload = 'auto';
    }
}

// キーボードショートカット（オプション）
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