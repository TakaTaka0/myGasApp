//アクセストークンを取得  
var slackAccessToken = 'your legacy token'; 
function myFunction() {
   var slackApp = SlackApp.create(slackAccessToken);
   
  //対象チャンネルID(チャンネル名だとNGなので、IDを取得)
  var channelId = "channel Name";
  
  //投稿メッセージ
  var message = "I am bot who can make something efficiently";
  
  var options = {
    username: "Takao_Hoshino"
  }
  slackApp.postMessage(channelId, message, options);
 }
