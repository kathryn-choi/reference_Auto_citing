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
    // var add_button      = $(".add_field_button"); //Add button ID
    
    // var x = 1; //initlal text box count
    // $(add_button).click(function(e){ //on add input button click
    //   e.preventDefault();
      
    //       x++; //text box increment
    //       input_id_list.push(x);
    //       $(wrapper).append('<div><br> name &nbsp&nbsp&nbsp&nbsp: <input type = "text" id = "name_'+x.toString()+'"><br>author&nbsp&nbsp&nbsp : <input type = "text" id = "author_'+x.toString()+'"><br>keyword : \
    //                         <input type = "text" id = "keyword_'+x.toString()+'"><a href="#" class="remove_field"> x </a></div>'); //add input box
      
    // });
    
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
    document.getElementById("rearrange").onclick = test;
  }
});
var x = 1;
function education_fields() {
  x++; //text box increment
  input_id_list.push(x);
  var wrapper = $(".input_fields_wrap"); 
  $(wrapper).append(
    '<div><div class="form-group"><hr>Name : &nbsp;&nbsp;&nbsp;&nbsp; <input type = "text" id = "name_'+x.toString()+'"><br>\
    </div><div class="form-group">Author : &nbsp;&nbsp;&nbsp;<input type = "text" id = "author_'+x.toString()+'"><br>\
    </div>Keyword : <input type = "text" id = "keyword_'+x.toString()+'">\
   <a href="#" class="remove_field"> x </a><br></div>'); //add input box
      
}
function remove_education_fields(rid) {
	 $('.removeclass'+rid).remove();
}

function test(){
  Word.run(function(context) {
    // Insert your code here. For example:
    var documentBody = context.document.body;
    context.load(documentBody);
   
    var author = document.getElementById("author_1").value;
    if(author.length != 0) {
      for(var i =0; i<input_id_list.length; i++)
        input_dict[input_id_list[i]] = [i+1];
    }
    return context.sync()
    .then(function(){
        var text = documentBody.text;
        var index_list = [];
        for(var i=0; i<input_id_list.length; i++){
          var index = text.indexOf('['+(i+1).toString()+']')
          input_dict[input_id_list[i]].push(index);
          index_list.push(index);
        }
        index_list.sort(function(a,b){
          return a-b;
        })
        for(var i=0; i<index_list.length; i++){
          for(var j=0; j<input_id_list.length; j++){
            if(input_dict[input_id_list[j]][1] == index_list[i]){
              if(input_dict[input_id_list[j]][0] != i+1){
                replace_num(input_dict[input_id_list[j]][0], (i+1));
              }
            }
          }
        }
    })
  });
}

function replace_text_multi(){
  for(var i =0; i<input_id_list.length; i++){
    var author = document.getElementById("author_"+input_id_list[i].toString()).value;
    // var name = document.getElementById("name_"+input_id_list[i].toString()).value;
    if(author.length != 0) {
      replace_text(author,i);
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
}

function replace_text(text,index){
  Word.run(function (ctx) {
    // Queue a command to search the document for the string "Contoso".
    // Create a proxy search results collection object.
    var results = ctx.document.body.search(text, {matchWholeWord:true});      //Search for the text to replace
    
    // Queue a command to load all of the properties on the search results collection object.
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

function replace_num(from_num, to_num){
  Word.run(function (ctx) {
    // Queue a command to search the document for the string "Contoso".
    // Create a proxy search results collection object.
    var results = ctx.document.body.search('['+from_num.toString()+']');      //Search for the text to replace
    
    // Queue a command to load all of the properties on the search results collection object.
    ctx.load(results, 'range');

    // Synchronize the document state by executing the queued commands,
    // and returning a promise to indicate task completion.
    return ctx.sync().then(function () {
      for (var i = 0; i < results.items.length; i++) {
        results.items[i].insertText('['+to_num.toString()+']',"replace");
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