WSL2とDocker desktopを使用します。

## 環境構築
### 1. WSL2の準備
**Windows Terminalから**以下のコマンドを実行してください。  

WSL2のインストール  
```$ wsl --install```  

WSL2を実行したときに使用するLinuxディストリビューションの設定  
使用可能なディストリビューションの表示  
```$ wsl -l -o```  
使用したいディストリビューションのインストール(今回はUbuntu22.04を使用しています)  
```$ wsl --install <DistroName>```   
ディストリビューションをデフォルトに設定(任意)  
```$ wsl -s <DistroName>```  

WSL2ターミナルの起動  
ディストリビューションを指定して起動するコマンド  
```$ wsl -d <DistroName>```  
デフォルトのLinuxを起動するコマンド(デフォルトに設定した場合)  
```$ wsl```  

### 2. Docker Desktop
以下のリンクからインストールをよく読み、インストールしてください。 
https://www.docker.com/ja-jp/get-started/  

WSL2 Backendの有効化  
Docker DesktopのSettings>Resauces>WSL integrationから以下の項目にチェックを入れてください。  
- [x] Enable integration with my default WSL distro  

またEnable integration with additional distrosにおける**Ubuntu-22.04の項目もON**にしてください。  

### 3. Docker Build  
Dockerをビルドするために以下のコマンドを**WSLターミナルで**実行してください。  
```
$ cd RECEIPT2EXCEL/docker
$ docker compose build
```

## 使用方法
以下のコマンドを**WSLターミナルで**実行してください。  
```$ docker compose up```  

ブラウザ```http://localhost:5137```にアクセスしてください。

または、  

Boot‗R2E.ps1内のパスを正しく書き換える。(初回のみ)  
Boot‗R2E.batをデスクトップに置き、ダブルクリックする。  