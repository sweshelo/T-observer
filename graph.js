/** 
 * Helper function to call MS Graph API endpoint
 * using the authorization bearer token scheme
*/

function callMSGraph(endpoint, method, token, body, callback) {

    const headers = new Headers();
    const bearer = `Bearer ${token}`;
    headers.append("Authorization", bearer);
    headers.append("Content-Type", "application/json");

    const options = {
        method: method,
        headers: headers,
    };

    if (body != null){
        options.body = JSON.stringify(body);
    }

    console.log(options);
    const req = new Request(endpoint, options);
    console.log('request made to Graph API at: ' + new Date().toString());

    fetch(req)
        .then(response => response.json())
        .then(response => callback(response, endpoint))
        .catch(error => console.log(error));
}

function sendMessage(message) {
    getTokenPopup(tokenRequest)
        .then(response => {
            let body = {
                "body":{
                    "contentType":"html",
                    "content": message,
                }
            };
            callMSGraph(graphConfig.graphTeamsTestChannelsEndpoint, "POST", response.accessToken, body,  ()=>{
                console.log(response);
            });
        }).catch(error => {
            console.error(error);
        });
}

function getPlans(callBack){
    getTokenPopup(tokenRequest)
        .then(response => {
            callMSGraph(graphConfig.graphPlannerEndpoint, "GET", response.accessToken, null, (e)=>{
                pObj = e.value;
                if (callBack !== null) {
                    callBack(e.value);
                }
            })
        });
}

function generateAlert(planObj){

    let allTask = 0, openingTask = 0, notouchTask = 0;
    let message = "";

    //メッセージテンプレ
    let WARN = [
        "<span style=\"background-color:#fcd116;margin-left:10px;padding:3px;\">⚠未完了</span> : ",
        "<span style=\"background-color:#ef6950;margin-left:10px;padding:3px;\">⚠未割り当て</span>",
    ]

    planObj.forEach((task)=>{
        //未完了
        if(task.completedBy == null){
            openingTask++;
            message += WARN[0] + task.title;

            //未割り当て
            if(!Object.keys(task.assignments).length){
                notouchTask++;
                message += WARN[1];
            }

            message += "<br />";
        }
    })

    //送信
    if (message != ""){
        message = "未完了のタスクが " +openingTask+ "件 あります。 <br><br>" + message;
        sendMessage(message);
    }
}


function check(){
    //clientId = document.getElementById("clientId").value;
    //authority = document.getElementById("authority").value;
    getPlans(()=>{
        let e = document.getElementById("log");
        let count = Object.keys(pObj).length;

        if (count > 0){
            e.value += "未完了のプランが存在します。\n\n";
            pObj.forEach(t =>{
                e.value += t.title+'\n';
            })
        }
    });
}

var pObj = null;

function post(){
    generateAlert(pObj);
}
