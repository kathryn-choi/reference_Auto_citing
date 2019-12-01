/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
var input_id_list = JSON.parse(localStorage.myIndexData);
var input_dict = JSON.parse(localStorage.myDicData);
var ref_array = JSON.parse(localStorage.myArrData);


Office.initialize = function () {
  if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
    console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    }
    document.getElementById("plus_paper").onclick = add_paper;
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
};

function add_paper(){
  Word.run(function (context) {

    var x = input_id_list.length;
    x++;
    console.log(x);
    input_id_list.push(x);
  
    var author = document.getElementById("author_1").value;
    var name = document.getElementById("name_1").value;
    var author_list = author.split(',');
    var publised_in = document.getElementById("publish_1").value;
    var pages = document.getElementById("page_1").value;
    var year = document.getElementById("year_1").value;
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
    //docBody.insertParagraph(ref_string,"End");
    ref_array.push(ref_string) ;
    localStorage.myArrData=JSON.stringify(ref_array);

    var i = x-1;
    if(author.length != 0) {
      replace_text(author,i);
      input_dict[input_id_list[i]] = [i+1];
     }
     
     // replace_text(name,i)
    var keyword = document.getElementById("keyword_1").value;
    if(keyword.length != 0){
     var keyword_list = keyword.split(',')
     for(var j=0; j<keyword_list.length; j++){
          replace_text(keyword_list[j],i)
        }
    }
    
    localStorage.myIndexData=JSON.stringify(input_id_list);
    localStorage.myDicData=JSON.stringify(input_dict);
    location.href="javascript:history.back()"

      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function insertParagraph() {
  Word.run(function (context) {
    console.log("start_insertParagraph")
    var docBody = context.document.body;
    docBody.insertParagraph("Office has several versions, including Office 2016, Office 365 Click-to-Run, and Office on the web.",
                            "Start");
      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
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