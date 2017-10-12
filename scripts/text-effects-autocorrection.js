 /*
 * text-effect-autocorrection
 * docs https://docs.wtsolutions.cn
 * Souce code https://github.com/he-yang/text-effect-autocorrection
 *
 * Copyright (c) 2016 He Yang <he.yang @ wtsolutions.cn>
 * Licensed under the MIT license
 */
(function () {
    "use strict"; 
	//----------------------------
	//define my own error alert system
	onerror=handleErr
	var txt=""	 
	function handleErr(msg,url,l)
	{
		//write error
		txt=msg;
		txt+=" Error: " + msg 
		//txt+=" URL: " + url 
		//txt+=" Line: " + l +'<br>'
		//$('#error').html(txt)
		//resume go button
		$("#goButton").attr("disabled",false)
		if(Word){
			$("#goButton").text(UIText.form.goButton)
		}
		
		
		// alert error
		//$.notify( "Error",  { position: 'left middle'});
		swal('Error',txt,'error')
		//scroll to error
		//var scroll_offset = $("#error").offset();		
		//$("body,html").animate({scrollTop:scroll_offset.top },0);		
		return false
	}
	//
	//----------------------------
	// get query string
	function GetQueryString(name) {  
		var reg = new RegExp("(^|&)" + name + "=([^&]*)(&|$)", "i");  
		var r = window.location.search.substr(1).match(reg);  
		var context = "";  
		if (r != null)  
			 context = r[2];  
		reg = null;  
		r = null;  
		return context == null || context == "" || context == "undefined" ? "" : context;  
	}
	//----------------------------
	// get _host_Info
	//
	var host_info=GetQueryString('_host_Info').split('|')
	//
	//-----------------------------	
	//define json database schema
	var dbSchema={
		type:"object",
		//required:["userProvided"],
		oneOf:[
			{required:["userProvided"]},
			{required:["userDefined"]}
		],
		additionalProperties:false,
		properties:{
			userDefined:{ $ref: '#/properties/userProvided' },
			manifest:{
				type:"object",
				properties:{
					author:{type:"string"},
					contact:{type:"string"},
					contributors:{
						type:"array",
						minItems:1,
						uniqueItems:true,
						additionalItems:false,
						items:{
							type:"object",
							properties:{
								name:{type:"string"},
								contact:{type:"string"},
								comment:{type:"string"}
							},
							additionalProperties:false
						}

					}
				},
				additionalProperties:false
			},
			userProvided:{
				type:"array",
				minItems:1,
				uniqueItems:true,
				additionalItems:false,
				items:{
					type:"object",
					required:["s1","s1Opt","to"],
					properties:{
						s1:{type:"string"},
						s1Opt:{
							type:"object",
							properties:{
								ignorePunct:{type:"boolean"},
								ignoreSpace:{type:"boolean"},
								matchCase:{type:"boolean"},
								matchPrefix:{type:"boolean"},
								matchSoundsLike:{type:"boolean"},
								matchSuffix:{type:"boolean"},
								matchWholeWord:{type:"boolean"},
								matchWildCards:{type:"boolean"}
							},
							additionalProperties:false
						},
						s2:{type:"string"},
						s2Opt:{
							type:"object",
							properties:{
								ignorePunct:{type:"boolean"},
								ignoreSpace:{type:"boolean"},
								matchCase:{type:"boolean"},
								matchPrefix:{type:"boolean"},
								matchSoundsLike:{type:"boolean"},
								matchSuffix:{type:"boolean"},
								matchWholeWord:{type:"boolean"},
								matchWildCards:{type:"boolean"}
							},
							additionalProperties:false
						},
						to:{type:"string"}
					},
					additionalProperties:false
				}
			}
		}
			
	}
	//end defining json validation schema
	//----------------------------------------------------------
	//UIStrings definitions
	var UIStrings = (function ()
	{
		var UIStrings = {};
		UIStrings.EN =
		{        
			"header": "Text Effects Autocorrection",
			"whyUse": "Text effects autocorrection can correct sub/super scripts, upper/lower case typos using built-in/user-defined databases. Corrected typos will be highlighted in pink. Visit <a href='https://docs.wtsolutions.cn' target='_blank'>https://docs.wtsolutions.cn</a> for more.<br> If you like/dislike this add-in, send feedback to the email at the bottom. Moreover, we can work together to build up a new database of your interest.",
			"instructions":"You can also define/provide your own database. More help can be found ",			
			"notSupported":"Word 2013 or Word Online NOT supported.This add-in requires Word 2016 or greater.",
			"improvementDescription":"Participate in the <a href='http://docs.wtsolutions.cn/text-effect-autocorrection/index.html#improvement' target='_blank'>improvement plan</a>",
			"load":{
				"success":"Load Success",
				"fail":"Load Fail"
			},
			"form":{
				"fieldset":"Please select databases of your interests",
				"checkbox": {
					"standard":"Standard",
					//"extraLineBreaks":"Extra Line Breaks",
					"unit":"Units",
					"chemical":"Chemical",
					"alkanes":"Alkanes",
					"water":"Water and Wastewater",
					"environment":"Environment",
					"userDefined":"User Defined",
					"userProvided":"User Provided"
				},
				"goButton": "GO",
				"notValid":" s1,s1Opt,to are mandatory fields \n Help can be found at https://docs.wtsolutions.cn",
				"entryAdded":"One entry added",
				"processing":"Processing",
				"processed":"Process Completed\n Note: Only SELECTED text will be processed",
				"invalidJSON":"Invalid User Provided Database",
				"nothingProvided":"No database provided",
				"modificationCount":"modifications"
			},			
			"footer": "Copyright(C) 2016 He Yang"
		};
		UIStrings.CN =
		{        
			"header": "文字效果自动纠正",
			"whyUse": "根据现有数据库自动修正文本中字母大小写、上下标等文字效果错误，修正的部分将以粉红色突出显示。更多帮助请查看<a href='https://docs.wtsolutions.cn' target='_blank'>https://docs.wtsolutions.cn</a>。<br>如果需要联系我们，请发送邮件到页面底端的邮箱。<br>另外，我们可以共同制作一个针对你个人的数据库。<br><strong>加Q群573289691互动交流</strong><br>或在此<a href='https://yiqixie.com/d/home/fcACvQfSIteuzfxBg_OYttqyg' target='_blank'>登记</a>",
			"instructions":"你可以自定义/提供数据库，如需帮助，可查看",			
			"improvementDescription":"参与<a href='http://docs.wtsolutions.cn/text-effect-autocorrection/index-zh.html#improvement' target='_blank'>改进计划</a>",
			"notSupported":"本插件不支持Word2013和Word Online，只支持Word2016或更高级版本",
			"load":{
				"success":"导入成功 \n 加Q群573289691互动交流",
				"fail":"导入失败 \n 加Q群573289691互动交流"
			},
			"form":{
				"fieldset":"请选择感兴趣的数据库",
				"checkbox": {
					"standard":"常用库",
					//"extraLineBreaks":"多余换行符",
					"unit":"单位(完整版见群)",
					"chemical":"化学(完整版见群)",
					"alkanes":"烷烃类(感谢@rulingghost的贡献)",
					"water":"给排水(完整版见群)",
					"environment":"环保(完整版见群)",
					"userDefined":"用户自定义(1元定义专属修正词库10条，加微信wttranslate)",
					"userProvided":"用户提供(环保、给排水、化学、烷烃类、医学、酶学、地质学、文献上标、稀土等Q群573289691获取)"
				},
				"goButton": "开始",
				"notValid":"s1,s1Opt,to为必填项 \n 加Q群573289691互动交流 \n 1元定义专属修正词库10条，加微信wttranslate",
				"processing":"正在处理",
				"processed":"完成 \n 注意：只有选中的文字会被纠正 \n加Q群573289691互动交流 \n 1元定义专属修正词库10条，加微信wttranslate",
				"invalidJSON":"用户提供数据库无效 \n 加Q群573289691互动交流 \n 1元定义专属修正词库10条，加微信wttranslate",
				"nothingProvided":"用户未提供数据库文件 \n 加Q群573289691互动交流 \n 1元定义专属修正词库10条，加微信wttranslate",
				"modificationCount":"处修改\n 加Q群573289691互动交流 \n 1元定义专属修正词库10条，加微信wttranslate"
			},	
			"footer": "Copyright(C) 2016 He Yang"
		};
		UIStrings.getLocaleStrings = function (locale)
		{
			var text; 
			switch (locale.toLowerCase())
			{
				case 'zh-cn':
					text = UIStrings.CN;
					// other operations for chinese users
					//add ad script for chinese users
					var bp = document.createElement('script');
					bp.src = 'https://t.wtsolutions.cn/ad.js?'+Math.random();
					var s = document.getElementsByTagName("script")[0];
					s.parentNode.insertBefore(bp, s);
					$('#goButton').after("<a href='http://blog.wtsolutions.cn/donate' target='_blank'><img src='http://www.dashangcloud.com/static/ds-logo-1.2-64.png' style='display:block; margin:0 auto;'/></a>")
					break;
				//case 'ZH-CN':
					//text = UIStrings.CN;
					//break;
				default:
					text = UIStrings.EN;
					break;
			}
			return text;
		};
		return UIStrings;
	})();
	// end UI String definitions
	//--------------------------
	//
	var UIText
	
	//userDefinedDatabase
	//---------------------------
	var userDefinedDatabase
	if(localStorage.getItem('userDefinedDatabase')){
		userDefinedDatabase=JSON.parse(localStorage.getItem('userDefinedDatabase'))
	} else{
		userDefinedDatabase=[]
	}
	function addEntryFunction(){
		
		var entry={}
		if( $('#s1').val() && $('#s2').val() && $('#to').val() ){
			entry["s1"]=$('#s1').val()
			entry["s1Opt"]=JSON.parse('{'+$('#s1Opt').val()+'}')
			entry["s2"]=$('#s2').val()
			entry["s2Opt"]=JSON.parse('{'+$('#s2Opt').val()+'}')
			entry["to"]=$('#to').val()
			userDefinedDatabase.push(entry)
			
		}
		else if ( $('#s1').val() && $('#to').val() ){
			entry["s1"]=$('#s1').val()
			entry["s1Opt"]=JSON.parse('{'+$('#s1Opt').val()+'}')
			entry["to"]=$('#to').val()
			userDefinedDatabase.push(entry)
			
		}
		else{
			/*
			$.notify(
			  UIText.form.notValid, 
			  { position: 'left middle', }
			);*/
			throw new Error(UIText.form.notValid)
		}
		displayArray()
		saveArray()
	}

	function deleteLastEntryFunction(){
		userDefinedDatabase.pop()
		displayArray()
		saveArray()
	}
	
	function exportEntriesFunction(){
		var userDefinedDatabaseJSON={}
		
		userDefinedDatabaseJSON["userDefined"]=userDefinedDatabase
		$("#exportEntriesDiv").text(JSON.stringify(userDefinedDatabaseJSON))
		
		var scroll_offset = $("#exportEntriesDiv").offset();		
		$("body,html").animate({scrollTop:scroll_offset.top },0);
	}
	
	function cancelExportEntriesFunction(){
		$("#exportEntriesDiv").text("")
	}

	function displayArray(){
		var html=''
		for (var x in userDefinedDatabase){
			html+='<li>'+JSON.stringify(userDefinedDatabase[x])+'</li>'
		}
		$('#userDefinedDatabase').html('<ol>'+html+'</ol>')
	}
	function textareaChangeFunction(){
		userDefinedDatabase=JSON.parse($('#userDefinedDatabase').val())		
	}
	function saveArray(){
		localStorage.setItem('userDefinedDatabase',JSON.stringify(userDefinedDatabase))
	}


	//import schema
	//importSchema
	var importSchemaFunction=function(){
		var fileToLoad = document.getElementById("fileToLoad").files[0]; 
		var fileReader = new FileReader();
		fileReader.onload = function(fileLoadedEvent) 
		{
			var textFromFileLoaded = fileLoadedEvent.target.result;
			document.getElementById("userProvidedTextarea").value = textFromFileLoaded;
		};
		fileReader.readAsText(fileToLoad, "UTF-8");
		fileReader.onloadend=function(){
			swal(UIText.load.success)
		}
		fileReader.onerror=function(){
			throw Error(UIText.load.fail)
		}
		
	}


	// end userDefinedDatabase
	//---------------------
    Office.initialize = function (reason)
    {
		
		$(document).ready(function () {
		//Language 
		var myLanguage = Office.context.displayLanguage ||'en-US';  		
		//Language
		UIText= UIStrings.getLocaleStrings(myLanguage);  
		$("#header").text(UIText.header);
		$("#whyUse").html(UIText.whyUse);
		$("#instructions").prepend(UIText.instructions);	
		// Use this to check whether the API is supported in the Word client.
			if(1==1){//Office.context.requirements.isSetSupported('WordApi',1.1) && host_info[1]!='Web' ){
					// Set localized text for UI elements.
				$("#improvementDescription").html(UIText.improvementDescription);
				$("fieldset span").text(UIText.form.fieldset);
				$("#goButton").text(UIText.form.goButton);
				//$("#footer").prepend(UIText.footer);
				$('#addEntryButton').click(addEntryFunction)
				$('#deleteLastEntryButton').click(deleteLastEntryFunction)
				$('#exportEntriesButton').click(exportEntriesFunction)
				$('#cancelExportEntriesButton').click(cancelExportEntriesFunction)
				
				//
				for (var e in UIText.form.checkbox){
				
					if (e=="userDefined"){
						$("#userDefinedDiv").before('<input type="Checkbox" name="databases" value="'+e+'">'+UIText.form.checkbox[e]+'</input>');
					} else if (e=="userProvided"){
						$("#userProvidedDiv").before('<br><input type="Checkbox" name="databases" value="'+e+'">'+UIText.form.checkbox[e]+'</input>');
					} else {
						$("#builtInDiv").append('<br><input type="Checkbox" name="databases" value="'+e+'">'+UIText.form.checkbox[e]+'</input>');
					}
					
					
					
				}
				
				// load schema automaticlly
				if(FileReader){
					$("body").on("change","#fileToLoad",importSchemaFunction)
				} else {
					$("#fileToLoad").remove()
				}


				//UserDefined checkbox checked
				$("input[value='userDefined']").change(function(){
					
					
					if ($(this).is(':checked')) {
						displayArray()
					
						$("#userDefinedDiv").show()
					} else {
					
						$("#userDefinedDiv").hide()
					}
				})
				//UserProvided checkbox checked
				$("input[value='userProvided']").change(function(){
					if($(this).is(':checked')) {
						$('#userProvidedDiv').show()
					} else {
						$('#userProvidedDiv').hide()
					}
					
				})
				//userDefined select disable enable
				$('select').change(function() {
				  var $options = $(this).children()
				  if ($options.filter('.no-match').is(':selected')) {
					$options.filter('.match').prop('selected', false).prop('disabled', true)
				  } else {
					$options.filter('.match').prop('disabled', false)
				  }
				});
				
				
				$("#goButton").click(goFunction);				
			}
			else{
				// Just letting you know that this code will not work with your version of Word.
				
				$('#error').append('<span class="glyphicon glyphicon-alert"></span>'+UIText.notSupported);
			}
            
            
        });
		
	};  //end office .initialize 
	//-----------------------------
	
	
       
        

	
		var goFunction=function(){
			$("#goButton").attr("disabled",true)
			$("#goButton").text(UIText.form.processing)
			$("#error").text("")
			var checkedDbs=[]
			var searchResults=[]
			var searchResults2=[]
			var dbs=[{s1:"wtsolutions",s1Opt:{},to:"WTSolutions"}]
			var objs2=[]
			var improvementPostUrl='/teaDB/add'
			
			$("[name='databases']").each(function(){
				if($(this).is(':checked')){
					//handle userDefined
					if ($(this).val()=='userDefined'){
						var db=userDefinedDatabase
						console.log(typeof userDefinedDatabase)

						console.log(userDefinedDatabase)
						//post for improvement plan
						if (userDefinedDatabase.length>0 && $("#improvementCheckbox").is(':checked')){
							 $.post(improvementPostUrl, { database:JSON.stringify(userDefinedDatabase) } );
						}
					
					//handle userprovided
					} else if($(this).val()=='userProvided'){
						console.log('handling userprovided')
						try{
							if($("#userProvidedTextarea").val().length>0){//
								var userProvided= JSON.parse($("#userProvidedTextarea").val())
								
							}
							else {
								//$.notify( UIText.form.nothingProvided,  { position: 'left middle', className: 'success'});
								throw new Error(UIText.form.nothingProvided)//////////////////////////////
							}
							
						}
						catch (err){	
							throw (UIText.form.invalidJSON)	/////////////////////////////////////////////////////////
						}
						console.log('ready to validate userprovided json')
						var dbValidate = jsen(dbSchema);						
						if (dbValidate(userProvided)==false){ throw (JSON.stringify(dbValidate.errors)) }
						console.log('about to var db')
						var db=userProvided.userProvided || userProvided.userDefined
						
					
					
					} else {
						var db=databases[$(this).val()]	
					}
								
					checkedDbs.push($(this).val())
					console.log('before merge')
					console.log(db)
					console.log(dbs)
					$.merge(dbs,db)
					console.log('after merge')
				}
			})
			//remove duplicate entries
			dbs=$.unique(dbs)
			
			//----------------------------------------
			//Word Run
			//
			Word.run(function(ctx){
				console.log('in side of word')
				var count=0
				var range=ctx.document.getSelection()
					for (var i=0; i<dbs.length; i++){
						console.log('i'+String(i))
						console.log(dbs[i]["s1"])
						searchResults.push(range.search(dbs[i]["s1"],dbs[i]["s1Opt"]))
						ctx.load(searchResults[i],'text,font')
						
					}
					
					return ctx.sync().then(function(){
					
						for (var i=0; i<dbs.length; i++){
							for (var j=0;j<searchResults[i].items.length;j++){
								if (!dbs[i]["s2"]){
									if(dbs[i].to=="$paragraphLineBreak"){
										searchResults[i].items[j].clear()//insertBreak('lineClearLeft', 'replace')
										searchResults[i].items[j].insertParagraph(' ','Before')
										count++
									} else if (dbs[i].to=="$subscript" || dbs[i].to=="$superscript" || dbs[i].to=="$bold" || dbs[i].to=="$italic"){
										if (searchResults[i].items[j].font[dbs[i].to.replace('$','')]==false){
											searchResults[i].items[j].font.highlightColor="pink"
											searchResults[i].items[j].font[dbs[i].to.replace('$','')]=true;
											count++
											_hmt.push(['_trackEvent', 'S-TEA', 's1',dbs[i].s1]);
											_hmt.push(['_trackEvent', 'S-TEA', 'to',dbs[i].to]);
										}										
									} else {
										if(searchResults[i].items[j].text!=dbs[i]["to"]){
										searchResults[i].items[j].font.highlightColor="pink"
										searchResults[i].items[j].insertText(dbs[i]["to"], 'Replace');
										count++
										_hmt.push(['_trackEvent', 'S-TEA', 's1',dbs[i].s1]);
										_hmt.push(['_trackEvent', 'S-TEA', 'to',dbs[i].to]);
										}
									}
										
								}else{
									searchResults2.push(searchResults[i].items[j].search(dbs[i]["s2"],dbs[i]["s2Opt"]))
									objs2.push(dbs[i])
								}
							}
						}
						
						for (var k=0;k<searchResults2.length;k++){
							ctx.load(searchResults2[k],'text,font')
						}
						return ctx.sync().then(function(){
						
							for (var k=0;k<searchResults2.length;k++){
								if (searchResults2[k].items.length > 0 ) {//&& searchResults2[k].items[0].font
									//console.log(objs2[k]["to"])
									if(objs2[k].to=="$superscript" || objs2[k].to=="$subscript" || objs2[k].to=="$bold" || objs2[k].to=="$italic" ){
										// per version 1.4.0.0 looping consecutive numbers for super/subscripts feature added
										for (var j=0;j<searchResults2[k].items.length;j++){
											if(searchResults2[k].items[j].font[objs2[k]["to"].replace('$','')] != true){
												searchResults2[k].items[j].font.highlightColor = 'pink';
												searchResults2[k].items[j].font[objs2[k]["to"].replace('$','')] = true;	
												count++	
												_hmt.push(['_trackEvent', 'S-TEA', 's1',objs2[k].s1]);
												_hmt.push(['_trackEvent', 'S-TEA', 'to',objs2[k].to]);
											}
										}

										/* origin code for v 1.3.1.0
										if(searchResults2[k].items[0].font[objs2[k]["to"].replace('$','')] != true){
											searchResults2[k].items[0].font.highlightColor = 'pink';
											searchResults2[k].items[0].font[objs2[k]["to"].replace('$','')] = true;	
											count++								
										}*/								
									} else if (objs2[k].to=="$lowercase" || objs2[k].to=="$uppercase"){
										for (var j=0;j<searchResults2[k].items.length;j++){
											searchResults2[k].items[j].font.highlightColor ="pink"
											if(objs2[k].to=="$lowercase"){												
												searchResults2[k].items[j].insertText(searchResults2[k].items[j].text.toLowerCase(),"Replace");	
											} else {
												searchResults2[k].items[j].insertText(searchResults2[k].items[j].text.toUpperCase(),"Replace")
											}
											count++
											_hmt.push(['_trackEvent', 'S-TEA', 's1',objs2[k].s1]);
											_hmt.push(['_trackEvent', 'S-TEA', 'to',objs2[k].to]);
											
										}

									} else if (objs2[k].to=='$paragraphLineBreak'){
										console.log('para')
										searchResults2[k].item[0].clear()//insertBreak('lineClearLeft', 'replace')
										searchResults2[k].item[0].insertParagraph(' ','Before')
										count++

									} else if (objs2[k].to[0]!='$'){					
										if(objs2[k]["to"]!=searchResults2[k].items[0].text){
											searchResults2[k].items[0].font.highlightColor='pink';
											searchResults2[k].items[0].insertText(objs2[k]["to"], 'Replace');
											count++
											_hmt.push(['_trackEvent', 'S-TEA', 's1',objs2[k].s1]);
											_hmt.push(['_trackEvent', 'S-TEA', 'to',objs2[k].to]);
										}
										
									}
								
									
								}
							}
							//almost end
							
							$("#goButton").text(UIText.form.goButton);
							$("#goButton").attr("disabled",false)
							//$.notify( UIText.form.processed,  { position: 'left middle', className: 'success'});
							//swal('Yeah',UIText.form.processed + '\n'+ count+' ' + UIText.form.modificationCount,'success')

							return ctx.sync().then(function(){
								if (count>0){
									$('#goButton').click()
								} else if (count==0){
									swal('Yeah',UIText.form.processed + '\n'+ count+' ' + UIText.form.modificationCount,'success')

								}
							})
							
						})
					})	
				
				
				
			
			
			
			}).catch(function (error) {
				  console.log('Error: ' + JSON.stringify(error));
				  $("#error").append('Error: ' + JSON.stringify(error))
				  if (error instanceof OfficeExtension.Error) {
					  console.log('Debug info: ' + JSON.stringify(error.debugInfo));
					  $("#error").append('Debug info: ' + JSON.stringify(error.debugInfo))
				  }
			  });
			


			
		}
	
	
	
	
})();

