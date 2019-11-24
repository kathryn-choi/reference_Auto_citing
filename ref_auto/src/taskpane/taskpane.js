/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

var input_id_list = [1];

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    var wrapper   		= $(".input_fields_wrap"); //Fields wrapper
    var add_button      = $(".add_field_button"); //Add button ID
    
    var x = 1; //initlal text box count
    $(add_button).click(function(e){ //on add input button click
      e.preventDefault();
      
          x++; //text box increment
          input_id_list.push(x);
          $(wrapper).append('<div><br>author : <input type = "text" id = "author_'+x.toString()+'"><br> name : <input type = "text" id = "name_'+x.toString()+'"><br>keyword : \
                            <input type = "text" id = "keyword_'+x.toString()+'"><a href="#" class="remove_field"> x </a></div>'); //add input box
      
    });
    
    $(wrapper).on("click",".remove_field", function(e){ //user click on remove text
      console.log(e);
      var deleted = e.target.previousElementSibling.id
      deleted = deleted.split('_')[1];
      deleted = parseInt(deleted);
      const idx = input_id_list.indexOf(deleted); 
      if (idx > -1) input_id_list.splice(idx, 1)
      console.log(input_id_list)
      e.preventDefault(); $(this).parent('div').remove(); 
    })
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

function replace_text_multi(){
  for(var i =0; i<input_id_list.length; i++){
    var author = document.getElementById("author_"+input_id_list[i].toString()).value;
    // var name = document.getElementById("name_"+input_id_list[i].toString()).value;
    var keyword = document.getElementById("keyword_"+input_id_list[i].toString()).value;
    var keyword_list = keyword.split(',')
    replace_text(author,i)
    // replace_text(name,i)
    for(var j=0; j<keyword_list.length; j++){
      replace_text(keyword_list[j],i)
    }
  }
}
//m_objectPath.m_objectPathInfo.id  
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

