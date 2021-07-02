# T-observer
Microsoft Planner にて登録済みのPlanの着手状況をみてMicrosoft Teamsにメッセージを飛ばすアプリです。  
`T-observer`という名前は以前作ったZabbix監視スクリプト`Z-observer`とまあまあ似たコンセプトで作成したのでそうしました。  

# 使い方

`Config_tmp.js`を`Config.js`にリネームし、行頭に次の5行を追加します。

```
const clientId = "";
const authority = "https://login.microsoftonline.com/";
const teamID = "";
const channelID = "[hogehoge1234]@thread.skype";
const planID = "";
```

続いて上記の値を正確に埋めます。  
`ClientID`、`authority`は[Microsoft Azure アプリの登録](https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)からアプリを登録すると貰える値です。  
`authority`は`"https://login.microsoftonline.com/"`から始まるURLです。  
  
`teamID`、`channelID`はアラートを飛ばすチームとチャネルを設定します。  
`teamID`は、チームの3点リーダをクリックして出現する「チームへのリンクをコピー」から取得できるURLの中のパラメータ`groupID`がこれに相当します。  
`channelID`は[Web版のTeams](https://teams.microsoft.com/)にアクセスした際に確認できるURLの中のパラメータ`threadId`がこれに相当します。  
  
`planID`は[Web版のplanner](https://tasks.office.com/)にアクセスした際に確認できるURLの中のパラメータ`planID`がこれに相当します。  
  
値を埋めれば作業は概ね終了です。  
最後にlocalhost:3000にサーバを立ててファイルを置き、ページにアクセスすれば利用可能かと思います。  

