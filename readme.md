# 使い方
## 必要ソフトウェア 
- node.js (npm)
- python (pip)

## 依存関係のインストール
```
pip install pyautogui
node i
```
## 変数設定
以下を適宜入力して同じフォルダ内にdata.jsとして保存
```
const user = {
    id:'[Microsoft ID]',
    passwd:'[Microsoft Password]'
}

chromepath = 'Google Chromeの実行ファイルへのフルパス';

module.exports.user = user
module.exports.chromepath = chromepath
```
## 実行
```
npm start
```