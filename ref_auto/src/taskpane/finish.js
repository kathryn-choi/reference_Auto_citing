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
    finishedList();
};

function finishedList() {
    var wrapper = $(".input_fields_wrap");
    for(i=0; i<input_id_list.length; i++){
        for(j=0; j<input_id_list.length; j++){
            if(input_dict[input_id_list[j]][2] == i){
                if(ref_array[input_id_list[j]][1].length == 0){
                    $(wrapper).append(
                    '<div role="button" id="ref_'+(i+1).toString + '" class="ms-welcome__action ms-Button ms-Button--hero ms-font-l " style = "width:300px;">\
                    ['+(i+1).toString()+']<br>'+ref_array[input_id_list[j]][0]+ '\
                    </div>');
                }
                else{
                    $(wrapper).append(
                    '<div role="button" id="ref_'+(i+1).toString + '" class="ms-welcome__action ms-Button ms-Button--hero ms-font-l" style = "width:300px;">\
                    ['+(i+1).toString()+']<br>'+ref_array[input_id_list[j]][0] + '<br>Keyword : '+ref_array[input_id_list[j]][1] + '\
                </div>');
                }
            }
        }
    }
  }