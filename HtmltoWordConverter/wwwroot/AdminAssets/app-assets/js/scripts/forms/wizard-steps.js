/*=========================================================================================
    File Name: wizard-steps.js
    Description: wizard steps page specific js
    ----------------------------------------------------------------------------------------
    Item Name: Frest HTML Admin Template
    Version: 1.0
    Author: PIXINVENT
    Author URL: http://www.themeforest.net/user/pixinvent
==========================================================================================*/
//    Wizard tabs with icons setup
// ------------------------------
$(".wizard-horizontal").steps({
  headerTag: "h6",
  bodyTag: "fieldset",
  transitionEffect: "fade",
  titleTemplate: '<span class="step">#index#</span> #title#',
/*  labels: {
    finish: 'Submit'
  },
  onFinished: function (event, currentIndex) {
    alert("Form submitted.");
  }*/
});
//        vertical Wizard       //
// ------------------------------
$(".wizard-vertical").steps({
  headerTag: "h3",
  bodyTag: "fieldset",
  transitionEffect: "fade",
  enableAllSteps: true,
  stepsOrientation: "vertical",
/*  labels: {
    finish: 'Submit'
  },
  onFinished: function (event, currentIndex) {
    alert("Form submitted.");
  }*/
});


//       Validate steps wizard //
// -----------------------------
// Show form
var stepsValidation = $(".wizard-validation");
var form = stepsValidation.show();

stepsValidation.steps({
  headerTag: "h6",
  bodyTag: "fieldset",
  transitionEffect: "fade",
  titleTemplate: '<span class="step">#index#</span> #title#',
  labels: {
    finish: 'Submit'
  },
  onStepChanging: function (event, currentIndex, newIndex) {
    // Allways allow previous action even if the current form is not valid!
    if (currentIndex > newIndex) {
      return true;
    }
    form.validate().settings.ignore = ":disabled,:hidden";
    return form.valid();
  },
  onFinishing: function (event, currentIndex) {
    form.validate().settings.ignore = ":disabled";
    return form.valid();
  },
  onFinished: function (event, currentIndex) {
    alert("Submitted!");
  }
});

// Initialize validation
stepsValidation.validate({
  ignore: 'input[type=hidden]', // ignore hidden fields
  errorClass: 'danger',
  successClass: 'success',
  highlight: function (element, errorClass) {
    $(element).removeClass(errorClass);
  },
  unhighlight: function (element, errorClass) {
    $(element).removeClass(errorClass);
  },
  errorPlacement: function (error, element) {
    error.insertAfter(element);
  },
  rules: {
    email: {
      email: true
    }
  }
});
// live Icon color change on state change
$(document).ready(function () {
  $(".current").find(".step-icon").addClass("bx bx-time-five");
  $(".current").find(".fonticon-wrap .livicon-evo").updateLiviconEvo({
    strokeColor: '#AD0132'
  });
   
});
// Icon change on state
// if click on next button icon change
$(".actions [href='#next']").click(function () {
  $(".done").find(".step-icon").removeClass("bx bx-time-five").addClass("bx bx-check-circle");
  $(".current").find(".step-icon").removeClass("bx bx-check-circle").addClass("bx bx-time-five");
  // live icon color change on next button's on click
  $(".current").find(".fonticon-wrap .livicon-evo").updateLiviconEvo({
    strokeColor: '#AD0132'
  });
  $(".current").prev("li").find(".fonticon-wrap .livicon-evo").updateLiviconEvo({
    strokeColor: '#39DA8A'
  });
});
$(".actions [href='#previous']").click(function () {
  // live icon color change on next button's on click
  $(".current").find(".fonticon-wrap .livicon-evo").updateLiviconEvo({
    strokeColor: '#AD0132'
  });
  $(".current").next("li").find(".fonticon-wrap .livicon-evo").updateLiviconEvo({
    strokeColor: '#adb5bd'
  });
});
// if click on  submit   button icon change
/*$(".actions [href='#finish']").click(function () {
  $(".done").find(".step-icon").removeClass("bx-time-five").addClass("bx bx-check-circle");
  $(".last.current.done").find(".fonticon-wrap .livicon-evo").updateLiviconEvo({
    strokeColor: '#39DA8A'
  });
});*/
// add primary btn class
$('.actions a[role="menuitem"]').addClass("btn btn-primary");
$('.icon-tab [role="menuitem"]').addClass("glow");
$('.wizard-vertical [role="menuitem"]').removeClass("btn-primary").addClass("btn-light-primary");

var i = 0;
function DynamicForm2() {
    nearByDiv = document.getElementById('ServiceItemForm');
    var e = nearByDiv.children;
    var length = e.length;

    i = length;
    var newdiv = document.createElement('DIV');
    newdiv.className = 'row justify-content-between';
    var showImageId = "'showImage" + i + "'";
    var iconNameId = "'iconName" + i + "'";
    var newForm1 = '<div class="col-md-3 col-sm-12"> <div class="form-group" ><label> Name  </label><select name="ServiceDetailsItem[' + i + '].Name"  class="form-control"><option value = "Include" >Include</option ><option value="Exclude">Exclude</option><option value="Recommended Gear">Recommended Gear</option><option value="Rules">Rules</option></select ></div></div>';
    var newForm2 = '<div class="col-md-3 col-sm-12"><div class="form-group"><label> Title  </label><input type="text" name="ServiceDetailsItem[' + i + '].Title" value="" class="form-control" placeholder="Title" /></div></div>';
    var newForm3 = '<div class="col-md-3 col-sm-12"><div class="form-group"><label> Description  </label><input type="text" name="ServiceDetailsItem[' + i + '].Description" value="" class="form-control" placeholder="Description" /></div></div>';
    var newForm5 = '<div class="col-12"><img src="/Images/ChooseImage.jpg" id="showImage' + i + '"   class="float-lg-right" style="width:150px; height:150px;" /></div>';
    var newForm4 = '<div class="col-md-3 col-sm-12"><fieldset class="form-group"><label>Icon</label><div class="custom-file"><input type="file" name="Icon" value="" class="custom-file-input" id="inputGroupFile0' + i + '" onchange="document.getElementById(' + showImageId + ').src = window.URL.createObjectURL(this.files[0]);document.getElementById(' + iconNameId + ').value = this.files[0].name;fun(this.files[0])" placeholder="Icon" /><label class="custom-file-label" for="inputGroupFile01">Choose file</label></div></fieldset></div>';

    var newForm7 = '<input type="hidden" class="form-control" id="iconName' + i + '" name="ServiceDetailsItem[' + i + '].Icon"  />';

    var completeForm = newForm1 + newForm2 + newForm3 + newForm4 + newForm5  + newForm7 + "<br />";
    newdiv.innerHTML = completeForm;
    nearByDiv.appendChild(newdiv);
}
function RemoveForm2() {
    var nearByDiv = document.getElementById('ServiceItemForm');
    document.getElementById('ServiceItemForm').removeChild(nearByDiv.lastChild);
}

//var i = 0;
//function DynamicForm2() {
//    nearByDiv = document.getElementById('ServiceItemForm');
//    var e = nearByDiv.children;
//    var length = e.length;

//    i = length;
//    var newdiv = document.createElement('DIV');
//    newdiv.className = 'row justify-content-between';
//   // var newForm1 = '<div class="col-md-3 col-sm-12"> <div class="form-group" ><label> Name  </label> <input type="text" name="ServiceDetailsItem[' + i + '].Name" value="" class="form-control" placeholder="Name" /></div></div>';
//    var newForm1 = '<div class="col-md-3 col-sm-12"> <div class="form-group" ><label> Name  </label><select name="ServiceDetailsItem[' + i + '].Name"  class="form-control"><option value = "Include" >Include</option ><option value="Exclude">Exclude</option><option value="Recommended Gear">Recommended Gear</option><option value="Rules">Rules</option></select ></div></div>';
//    var newForm2 = '<div class="col-md-3 col-sm-12"><div class="form-group"><label> Title  </label><input type="text" name="ServiceDetailsItem[' + i + '].Title" value="" class="form-control" placeholder="Title" /></div></div>';
//    var newForm3 = '<div class="col-md-3 col-sm-12"><div class="form-group"><label> Description  </label><input type="text" name="ServiceDetailsItem[' + i + '].Description" value="" class="form-control" placeholder="Description" /></div></div>';
//    var newForm4 = '<div class="col-md-3 col-sm-12"><fieldset class="form-group"><label>Icon</label><div class="custom-file"><input type="file" name="Icon" value="" class="custom-file-input" id="inputGroupFile01" placeholder="Icon" /><label class="custom-file-label" for="inputGroupFile01">Choose file</label></div></fieldset></div>';

//       var completeForm = newForm1 + newForm2 + newForm3 + newForm4 + "<br />";
//    newdiv.innerHTML = completeForm;
//    nearByDiv.appendChild(newdiv);
//}
//function RemoveForm2() {
//    var nearByDiv = document.getElementById('ServiceItemForm');
//    document.getElementById('ServiceItemForm').removeChild(nearByDiv.lastChild);
//}

var a = 0;
function DynamicForm3() {
    nearByDiv = document.getElementById('ServiceItenieraryForm');
    var e = nearByDiv.children;
    var length = e.length;

    a = length;
    var newdiv = document.createElement('DIV');
    newdiv.className = 'row justify-content-between';
    var newForm1 = '<div class="col-md-6 col-sm-12"> <div class="form-group" ><label> Title  </label> <input type="text" name="ServiceItenierary[' + a + '].Title" value="" class="form-control" placeholder="Title" /></div></div>';
    var newForm2 = '<div class="col-md-6 col-sm-12"><div class="form-group"><label> Description  </label><input type="text" name="ServiceItenierary[' + a + '].Description" value="" class="form-control" placeholder="Description" /></div></div>';
     var completeForm = newForm1 + newForm2  + "<br />";
    newdiv.innerHTML = completeForm;
    nearByDiv.appendChild(newdiv);
}
function RemoveForm3() {
    var nearByDiv = document.getElementById('ServiceItenieraryForm');
    document.getElementById('ServiceItenieraryForm').removeChild(nearByDiv.lastChild);
}




var b = 0;
function DynamicForm4() {
    nearByDiv = document.getElementById('ServiceAddOnsForm');
    var e = nearByDiv.children;
    var length = e.length;

    b = length;
    var newdiv = document.createElement('DIV');
    newdiv.className = 'row justify-content-between';
    var newForm1 = '<div class="col-md-6 col-sm-12"> <div class="form-group" ><label> Name  </label> <input type="text" name="ServiceAddOns[' + b + '].ServiceAddOnsName" value="" class="form-control" placeholder="Name" /></div></div>';
    var newForm2 = '<div class="col-md-6 col-sm-12"><div class="form-group"><label> Price  </label><input type="text" name="ServiceAddOns[' + b + '].ServiceAddOnsPrice" value="" class="form-control" placeholder="Price" /></div></div>';
     var completeForm = newForm1 + newForm2  + "<br />";
    newdiv.innerHTML = completeForm;
    nearByDiv.appendChild(newdiv);
}
function RemoveForm4() {
    var nearByDiv = document.getElementById('ServiceAddOnsForm');
    document.getElementById('ServiceAddOnsForm').removeChild(nearByDiv.lastChild);
}


/*
var c = 0;
function DynamicForm5() {
    nearByDiv = document.getElementById('ServicePriceForm');
    var e = nearByDiv.children;
    var length = e.length;

    c = length;
    var newdiv = document.createElement('DIV');
    newdiv.className = 'row justify-content-between';
    var newForm1 = '<div class="col-md-4 col-sm-12"> <div class="form-group" ><label> AdultsPrice  </label> <input type="text" name="ServicePrice[' + c + '].AdultsPrice" value="" class="form-control" placeholder="AdultsPrice" /></div></div>';
    var newForm2 = '<div class="col-md-4 col-sm-12"><div class="form-group"><label> InfantsPrice  </label><input type="text" name="ServicePrice[' + c + '].InfantsPrice" value="" class="form-control" placeholder="InfantsPrice" /></div></div>';
    var newForm3 = '<div class="col-md-4 col-sm-12"><div class="form-group"><label> ChildPrice  </label><input type="text" name="ServicePrice[' + c + '].ChildPrice" value="" class="form-control" placeholder="ChildPrice" /></div></div>';
    var completeForm = newForm1 + newForm2 + newForm3  + "<br />";
    newdiv.innerHTML = completeForm;
    nearByDiv.appendChild(newdiv);
}
function RemoveForm5() {
    var nearByDiv = document.getElementById('ServicePriceForm');
    document.getElementById('ServicePriceForm').removeChild(nearByDiv.lastChild);
}*/




var d = 0;
function DynamicForm6() {
    nearByDiv = document.getElementById('ServiceDateForm');
    var e = nearByDiv.children;
    var length = e.length;

    d = length;
    var newdiv = document.createElement('DIV');
    newdiv.className = 'row justify-content-between';
    var newForm1 = '<div class="col-md-6 col-sm-12"> <div class="form-group" ><label> StartDate  </label> <input type="date" name="ServiceDate[' + d + '].StartDate" value="" class="form-control" placeholder="StartDate" /></div></div>';
    var newForm2 = '<div class="col-md-6 col-sm-12"><div class="form-group"><label> EndDate  </label><input type="date" name="ServiceDate[' + d + '].EndDate" value="" class="form-control" placeholder="EndDate" /></div></div>';
        var completeForm = newForm1 + newForm2  + "<br />";
    newdiv.innerHTML = completeForm;
    nearByDiv.appendChild(newdiv);
}
function RemoveForm6() {
    var nearByDiv = document.getElementById('ServiceDateForm');
    document.getElementById('ServiceDateForm').removeChild(nearByDiv.lastChild);
}


var f = 0;
function DynamicForm7() {
    nearByDiv = document.getElementById('ServiceImageForm');
    var e = nearByDiv.children;
    var length = e.length;

    f = length;
    var newdiv = document.createElement('DIV');
    newdiv.className = 'row justify-content-between';
    var showImageId = "'showImagee" + f + "'";
    var iconNameId = "'iconNamee" + f + "'";
    var newForm4 = '<div class="col-md-6 col-sm-12 "><fieldset class="form-group"><br /> <br /> <br /> <br /><label>ImagePath</label><div class="custom-file"><input type="file" name="file" value="" class="custom-file-input" id="inputGroupFile' + f + '" onchange="document.getElementById(' + showImageId + ').src = window.URL.createObjectURL(this.files[0]);document.getElementById(' + iconNameId + ').value = this.files[0].name;fun(this.files[0])" placeholder="Image" /><label class="custom-file-label" for="inputGroupFile0">Choose file</label></div></fieldset></div>';
    var newForm5 = '<div class="col-md-6 col-sm-12"><br /> <br /><img src="/Images/ChooseImage.jpg " id="showImagee' + f + '" class="float-lg-right" style="width:250px; height:150px;" /></div>';

    var newForm7 = '<input type="hidden" class="form-control" id="iconNamee' + f + '" name="ServiceImage[' + f + '].ImagePath"  />';
  
    var completeForm = newForm4 + newForm5 + newForm7  + "<br />";

   
    newdiv.innerHTML = completeForm;
    nearByDiv.appendChild(newdiv);
}
function RemoveForm7() {
    var nearByDiv = document.getElementById('ServiceImageForm');
    document.getElementById('ServiceImageForm').removeChild(nearByDiv.lastChild);
}

var g = 0;
function DynamicForm8() {
    nearByDiv = document.getElementById('PickipLocationForm');
    var e = nearByDiv.children;
    var length = e.length;

    g = length;
    var newdiv = document.createElement('DIV');
    newdiv.className = 'row justify-content-between';
    var newForm1 = '<div class="col-md-6 col-sm-12"> <div class="form-group" ><label> Name  </label> <input type="text" name="PickupLocation[' + g + '].Name" value="" class="form-control" placeholder="Name" /></div></div>';
    var newForm2 = '<div class="col-md-6 col-sm-12"><div class="form-group"><label> Price  </label><input type="text" name="PickupLocation[' + g + '].Price" value="" class="form-control" placeholder="Price" /></div></div>';
    var completeForm = newForm1 + newForm2 + "<br />";
    newdiv.innerHTML = completeForm;
    nearByDiv.appendChild(newdiv);
}
function RemoveForm8() {
    var nearByDiv = document.getElementById('PickipLocationForm');
    document.getElementById('PickipLocationForm').removeChild(nearByDiv.lastChild);
}
