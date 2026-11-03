
(function(){
 var timer;
 var boot=window.wlBoot={began:false,done:false,ready:function(){boot.done=true;window.clearTimeout(timer);document.getElementById('startup').hidden=true;},fail:function(title,message){if(boot.done)return;document.getElementById('startup-title').textContent=title;document.getElementById('startup-message').textContent=message;document.getElementById('startup-help').textContent='請資訊人員檢查網站連線。';}};
 window.addEventListener('error',function(){boot.fail('網站程式未能啟動','請使用新版 Edge 或 Chrome。');});
 timer=window.setTimeout(function(){if(boot.began)boot.fail('網站連線等待過久','請重新載入，或聯絡資訊人員。');else boot.fail('網站程式未能啟動','請使用新版 Edge 或 Chrome。');},10000);
})();
