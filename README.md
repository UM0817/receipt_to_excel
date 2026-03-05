WSL2とDocker desktopを使用します。

## 1. WSL2の準備
Windows Terminalから以下のコマンドを実行してください。  
### WSL2のインストール  
```$ wsl --install```  

### WSL2を実行したときに使用するLinuxディストリビューションの設定  
使用可能なディストリビューションの表示  
```$ wsl -l -o```  
使用したいディストリビューションのインストール(今回はUbuntu22.04を使用しています)  
```$ wsl --install <DistroName>```   
ディストリビューションをデフォルトに設定(任意)  
```$ wsl -s <DistroName>```  

### WSL2ターミナルの起動
デフォルトのLinuxを起動するコマンド  
```$ wsl```  
ディストリビューションを指定して起動するコマンド  
```$ wsl -d <DistroName>```  

## 2. Docker Desktop
以下のリンクからインストールしてください  
https://www.docker.com/ja-jp/get-started/  
WSL2 Backendの有効化  
Settings>Resauces>WSL integrationから以下の項目にチェックを入れてください  
- [x] Enable integration with my default WSL distro  

またEnable integration with additional distrosにおける**Ubuntu-22.04の項目もON**にしてください  

## 3. Docker Build  
Dockerをビルドするために以下のコマンドを**WSLターミナルで**実行してください  
```
$ cd RECEIPT2EXCEL/docker
$ docker compose build
$ docker compose up
```
