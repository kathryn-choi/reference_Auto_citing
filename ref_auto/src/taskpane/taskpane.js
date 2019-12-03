/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

var input_id_list = [1];
var input_dict = {};

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    var wrapper   		= $(".input_fields_wrap"); //Fields wrapper
    
    $(wrapper).on("click",".remove_field", function(e){ //user click on remove text
      console.log(e);
      var deleted = e.target.previousElementSibling.id
      deleted = deleted.split('_')[1];
      deleted = parseInt(deleted);
      const idx = input_id_list.indexOf(deleted); 
      if (idx > -1) input_id_list.splice(idx, 1)
      console.log(input_id_list)
      e.preventDefault(); $(this).parent('div').remove(); 
      e.preventDefault(); $(this).parent('div').remove(); 
      e.preventDefault(); $(this).parent('div').remove(); 
      e.preventDefault(); $(this).parent('div').remove(); 
    })

    document.getElementById("add").onclick = education_fields;

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
      console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("send").onclick = replace_text_multi;
  }
});
var x = 1;
function education_fields() {
  x++; //text box increment
  input_id_list.push(x);
  var wrapper = $(".input_fields_wrap"); 
  $(wrapper).append(
    '<div><div class="form-group"><hr>Name &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : &nbsp; <input type = "text" id = "name_'+x.toString()+'"><br>\
    </div><div class="form-group"> Author &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: &nbsp;<input type = "text" id = "author_'+x.toString()+'"><br>\
    </div><div class="form-group"> Published in &nbsp;: &nbsp;<input type = "text" id = "publish_'+x.toString()+'"><br>\
    </div><div class="form-group"> Pages   &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : &nbsp;<input type = "text" id = "page_'+x.toString()+'"><br>\
    </div><div class="form-group"> Year &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: &nbsp;<input type = "text" id = "year_'+x.toString()+'"><br>\
    </div> Keyword &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: &nbsp;<input type = "text" id = "keyword_'+x.toString()+'">\
   <a href="#" class="remove_field"> x </a><br></div>'); //add input box
      
}
function remove_education_fields(rid) {
	 $('.removeclass'+rid).remove();
}

function replace_text_multi(){
  for(var i =0; i<input_id_list.length; i++){
    var author = document.getElementById("author_"+input_id_list[i].toString()).value;
    // var name = document.getElementById("name_"+input_id_list[i].toString()).value;
    if(author.length != 0) {
      var author_list = author.split(',')
      for(var j =0; j<author_list.length; j++){
        replace_text(author_list[j],i);
      }
      input_dict[input_id_list[i]] = [i+1];
    }
    // replace_text(name,i)
    var keyword = document.getElementById("keyword_"+input_id_list[i].toString()).value;
    if(keyword.length != 0){
      var keyword_list = keyword.split(',')
      for(var j=0; j<keyword_list.length; j++){
        replace_text(keyword_list[j],i)
      }
    }
  }
  localStorage.myIndexData=JSON.stringify(input_id_list);
  localStorage.myDicData=JSON.stringify(input_dict);
  insert_references();
  location.href="src/taskpane/afterinsert.html";
}

function replace_text(text,index){
  Word.run(function (ctx) {
    var results = ctx.document.body.search(text, {matchWholeWord:true});      //Search for the text to replace
    
    ctx.load(results, 'range');

    // Synchronize the document state by executing the queued commands,
    // and returning a promise to indicate task completion.
    return ctx.sync().then(function () {
      for (var i = 0; i < results.items.length; i++) {
        results.items[i].insertText("["+(index+1).toString()+"]", "After");
      }
    })
    // Synchronize the document state by executing the queued commands.
    .then(ctx.sync)
    .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });

  });
}



function insert_references(){
  
  Word.run(function (context) {
    var ref_array = {};
    for(var i =0; i<input_id_list.length; i++){
      var author = document.getElementById("author_"+input_id_list[i].toString()).value;
      var name = document.getElementById("name_"+input_id_list[i].toString()).value;
      var author_list = author.split(',');
      var publised_in = document.getElementById("publish_"+input_id_list[i].toString()).value;
      var pages = document.getElementById("page_"+input_id_list[i].toString()).value;
      var year = document.getElementById("year_"+input_id_list[i].toString()).value;
      var keyword = document.getElementById("keyword_"+input_id_list[i].toString()).value;
      var ref_string = "";
      
      for(var j=0; j<author_list.length; j++){
        ref_string += author_list[j];
        if((author_list.length != 1) && j == author_list.length -2){
          ref_string += " and ";
        }
        else{
          ref_string += ", "
        }
      } 
      
      ref_string += '"' + name + '," in ' + publised_in + ", " + year + ". pp. " + pages +"."
      ref_array[input_id_list[i]] = [ref_string] ;
      ref_array[input_id_list[i]].push(keyword);
      localStorage.myArrData=JSON.stringify(ref_array);
    }

    return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}
