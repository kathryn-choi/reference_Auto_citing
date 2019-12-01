/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
var input_id_list = JSON.parse(localStorage.myIndexData);
var input_dict = JSON.parse(localStorage.myDicData);
var ref_array = JSON.parse(localStorage.myArrData);

Office.initialize = function () {

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
      console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("rearrange").onclick = test;
    insertRefList();
};
var x = 1;


function insertRefList() {
    Word.run(function (context) {
        var wrapper = $(".input_fields_wrap");
        $(wrapper).empty();
        for(var i = 0; i < ref_array.length; i++)
        {
            //docBody.insertParagraph(ref_array[i],"End");
            $(wrapper).append(
                '<div role="button" id="ref_'+(i+1).toString + '" class="ms-welcome__action ms-Button ms-Button--hero ms-font-xl" onclick ="insert_num(\'' + (i+1).toString() + '\')">\
                <span class="ms-Button-label">'+ref_array[i]+'</span>\
            </div>');
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

function test(){
    Word.run(function(context) {
      // Insert your code here. For example:
      var documentBody = context.document.body;
      context.load(documentBody);
     
      //var author = document.getElementById("author_1").value;
      if(ref_array.length != 0) {
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
                input_dict[input_id_list[j]].push(i);
                if(input_dict[input_id_list[j]][0] != i+1){
                  replace_num(input_dict[input_id_list[j]][0], (i+1));
                }
              }
            }
          }
          for(i=0; i<input_id_list.length; i++){
            for(j=0; j<input_id_list.length; j++){
               if(input_dict[input_id_list[j]][2] == i){
                documentBody.insertParagraph("["+(i+1).toString() +"] " + ref_array[j],"End");
               }
             }
           }
      })
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

  function insert_num(index) {
    Word.run(function (context) {
  
      
      var doc = context.document;
      var originalRange = doc.getSelection();
      console.log(index);
      var it = "["+ index +"] "
      originalRange.insertText(it, "After");
      
      originalRange.load("text");
      return context.sync()
        .then(context.sync);
    })
    .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });  
}