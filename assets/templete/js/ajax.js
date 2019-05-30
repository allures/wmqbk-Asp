function ckradd(e,f){
  if($("#"+e+"pname").val()==""){
  errmsg("请输入昵称后再提交");
  $("#"+e+"pname").focus();
  return false;
  }
  var val=$("#"+e+"plog").val();
  if(val.length<5 || val.length>200){
	errmsg("评论内容必须在5-200字之间，请修改后再提交！");
   $("#"+e+"plog").focus();
   return false;  	  
  }
  var code=$("#safecode").val();
  if (f=='1'&&code==''){	
      errmsg("请正确输入右侧答案！");$("#safecode").focus();return false; 
 } 
 return true
}

function ckse(){
	var val=$("#key").val();
  if(val.length<2 || val.length>10){
	alert("关键词必须在2-10字之间，请修改后再提交！");
   $("#key").focus();
   return false;  	  
  }
}
function StopButton(id,s){
	$("#"+id).attr("disabled",true);　
	$("#"+id).text("提交("+s+")");
	if(--s>0){
		 setTimeout("StopButton('"+id+"',"+s+")",1000);
	}
	if(s<=0){
		$("#"+id).text(' 提 交 ');
	    $("#"+id).removeAttr("disabled");
	} 
}
 
function errmsg(s,el=''){
   $(el+'#errmsg').show().text(s).fadeOut(2000);
}
function savelog() {
	var tit=$("#tit").val(),
		summ = $("#sum").val(),
		log = ndPanel.nicInstances[0].getContent(),
		pic = upic,
		pics = pic_arr.join(','),
		id = $("#id").val(),
		c = $("#c").val(),
		pass=$("#pass").val(),
		atime = $("#atime").val();
	    hide = $('#hide').prop("checked")?1:0;
		lock = $('#lock').prop("checked")?1:0;
    if(log =="" ){
      errmsg("写点什么吧！");
      $("#log").focus();
      return false;
  }
	$.post("./app/class/ajax.asp?act=savelog&id=" + id, {
		tit:tit,
		summ:summ,
		logs: log,
		pic: pic,
		pics: pics,
		atime:atime,
		pass:pass,
		hide:hide,
		lock:lock,
		c: c
	}, function(data) {
		if (data.result == '200') {	 ;
		  window.location.href = window.location.pathname+'?act=pl&id='+data.message;
		}else{
		   errmsg(data.message)
		}
	}, 'json');

}
function saveset(){
    var data = $("#formset").serialize();
    $.post("./app/class/ajax.asp?act=saveset",data , function(data) {
		errmsg(data.message)
	}, 'json');

}
function savewid(id){
   var data = $("#formwid"+id).serialize();
   $.post("./app/class/ajax.asp?act=savewid&id="+id,data , function(data) {
		errmsg(data.message,"#formwid"+id+' ')
	}, 'json');

}
function delwid(id){
	if(confirm('确定要删除吗?'))
	{	
		$.get("./app/class/ajax.asp?act=delwid&id="+id,function(data){if(data.result=='200'){ window.location.reload();}else{alert(data.message);}},'json');
     }
 }

function dellog(id,v){
	if(confirm('确定要删除吗?'))
	{	
		$.get("./app/class/ajax.asp?act=dellog&id="+id,function(data){if(data.result=='200'){ if(v=='1'){location.href="./";}else{$("#log-"+id).fadeOut();} }else{alert(data.message);}},'json');
     }
}
function delpl(id,pid){
	if(confirm('确定要删除吗?'))
	{	
		$.get("./app/class/ajax.asp?act=delpl&id="+id+"&pid="+pid,function(data){if(data.result=='200'){$("#Com-"+pid).fadeOut();}else{alert(s);}},'json');
     }
}
function shpl(id){
		$.get("./app/class/ajax.asp?act=shpl&pid="+id,function(data){if(data.result=='200'){$("#sh-"+id).fadeOut();}else{alert(data.message);}},'json');
}
function zdlog(id){
	var zdobj=$("#zd-"+id);
	var xval=0;
	if(zdobj.text()=='置顶'){xval=1};
	$.get("./app/class/ajax.asp?act=zdlog&id="+id+"&x="+xval,function(data){if(data.result=='200'){zdobj.text(data.message);}else{alert(data.message);}},'json');
}
function addpl(id,f){	
	var ck = ckradd('',f);
	if (ck ===false)
	{
		return ck;
	}
	var npname = $("#pname").val(),nplog = $("#plog").val(),nscode=$("#safecode").val();
	$.post("./app/class/ajax.asp?act=addpl&id="+id, {pname:npname, plog:nplog,scode:nscode}, function(txt) {
	 if(txt.result==500){alert(txt.message);$("#safecode").val('');reloadcode();$("#safecode").focus();}else
	 {$("#comment_list").append(txt.message);$("#plog").val('');$("#safecode").val('');reloadcode();StopButton('add',9);}											 
	},'json');		
}
function repl(pid,cid){
	var ore = $('#Ctext-'+pid).find('.re span').text();
	var x = 1;
	if (ore == ""){x=0;}
    var rebox = '<div class="rebox"><input placeholder="随便说点什么吧..." name="rlog" rows="3" id="rlog" class="log relog" value="'+ore+'"> <button name="re" id="re" class="btn" onclick="plsave('+cid+','+pid+','+x+')"> 回 复 </button> <button onclick="capl()" class="btn"> 取 消 </button></div>';
	$('.rebox').remove();
	$('#Ctext-'+pid).append(rebox);
}
function capl(){
	$('.rebox').remove();
}
function plsave(id,pid,x){	
	var rlog = $("#rlog").val();
	if(rlog==''){
		alert('请输入内容后再提交');
		return false;
	}
	$.post("./app/class/ajax.asp?act=plsave&id=" + id + "&pid=" + pid, {
		rlog: rlog
	}, function(data) {
        capl();
		if (data.result == '200') {
			if(x==1){
				$('#Ctext-'+pid).find('.re span').text(data.message);
			}else{
				$('#Ctext-'+pid).append('<p class="re">&nbsp;&nbsp;<strong style="color:#C00">回复</strong>：<span>'+data.message+'</span></p>');
			}
			 
		} else {
			alert(data.message);
		}
	}, 'json');}
function ckpass(id){	
	var ps= $("#password").val();
	if (ps!=''){
	$.post("./app/class/ajax.asp?act=ckpass&id="+id, {ps:ps}, function(data) {if(data.result=='200'){ $("#password").parent().html(data.message)}else{alert(data.message);}},'json');}	
}
function upCache(){
   $.get("./app/class/ajax.asp?act=upcache",function(data){alert(data.message);},'json');
}
function DotRoll(elm) {
    $("body,html").animate({ scrollTop: $("a[name='" + elm + "']").offset().top }, 500);
}
function reloadcode(){$('#codeimg').attr('src','./app/class/codes.asp?n='+Math.random());}
function getFileName(o){
    var pos=o.lastIndexOf("\\");
    return o.substring(pos+1);  
}
function up_callback(p){
     pic_arr.push(p);
 }
function showImg(input) {
        var img = input.files[0];
		    if(!(img.type.indexOf('image')==0 && img.type && /\.(?:jpeg|jpg|png|gif)$/.test(img.name)) ){
                alert('图片只能是jpg,gif,png格式');
                return ;
            } 
		resize(img, 180, 120, upload)
}

function upload(imgdata){
            $.post("app/class/upload.asp?act=pic&pic="+upic, {img: imgdata}, function(ret){
				upic = ret.url
			    imgdata==''?$('#delpic').hide():$('#delpic').show()
            }, 'json');
}

function resize(img, width, height, callback){
            // 创建临时图片对象
            var image = new Image;
            // 创建画布
          
			try {
                   var canvas = document.createElement('canvas');
                   var context = canvas.getContext('2d');
             } catch(e) {
				 alert('浏览器不支持canvas,请使用html5浏览器');
                 return
              }
 
            // 临时图片加载
            image.onload = function(){
                // 图片尺寸
                var img_w = image.naturalWidth; 
                var img_h = image.naturalHeight;
                // 缩略后尺寸
                var dimg_w;
                var dimg_h;
			    var n = Math.max((width/img_w),(height/img_h));
				dimg_w = Math.ceil(img_w*n);
                dimg_h = Math.ceil(img_h*n);
                // 定义画布尺寸
                canvas.width = width;
                canvas.height = height; 
                context.drawImage(image, 0, 0, dimg_w, dimg_h); 
                // 获取画布数据
                var imgdata = canvas.toDataURL(img.type);
				$("#pic").attr("src",imgdata).show();
                // 将画布数据回调返回
                if(typeof(callback)==='function'){
                   callback(imgdata);
                }

            }
            // file reader
            var reader = new FileReader();
            reader.readAsDataURL(img);
            reader.onload = function(e){
                image.src = reader.result;
            }

        }