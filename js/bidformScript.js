// JavaScript for bid request form

// uncomment script tages if putting back in html
//<!-- Validation Script -->
// <script language="JavaScript">

// object definition
function validation(realName, formEltName, eltType, upToSnuff, format) {
    this.realName = realName;
    this.formEltName = formEltName;
    this.eltType = eltType;
    this.upToSnuff = upToSnuff;
    this.format = format;}

// create a new object for each form element to be validated
var project = new validation('Type of Project', 'Project', 'text', 'isSelect(formObj)', null);
var status = new validation('Design Status', 'Status', 'text', 'isSelect(formObj)', null);
var startdate = new validation('StartDate', 'StartDate', 'text', 'isSelect(formObj)', null);
var flexible = new validation('Your Timing is Flexible', 'TimingFlexible', 'radio', 'isRadio(formObj)', null);
var HaveNames = new validation('if you are working with other contractors', 'HaveNames', 'radio', 'isRadio(formObj)', null);
var HaveBid = new validation('if you have already received a bid', 'HaveBid', 'radio', 'isRadio(formObj)', null);
var description = new validation('Project Details', 'description', 'text', 'isText(str)', null);
var address = new validation('Project Adress', 'street', 'text', 'isText(str)', null);
var city = new validation('City', 'city', 'text', 'isText(str)', null);
var zip = new validation('Project Zip Code', 'zip', 'text', 'isText(str)', null);
var first = new validation('First Name', 'fname', 'text', 'isText(str)', null);
var last = new validation('Last Name', 'lname', 'text', 'isText(str)', null);
var phone = new validation('Phone Number', 'phone', 'text', 'isText(str)', null);
var besttime = new validation('Best Time to Call', 'besttime', 'text', 'isText(str)', null);

var elts = new Array(
               project,
               status,
			   startdate,
			   flexible,
               HaveNames,
               HaveBid,
			   description,
			   address,
			   city,
			   zip,
			   first,
			   last,
			   phone,
			   besttime);

var allAtOnce = false;

/***************************************************************
 ** change text for alerts here
 ** unspecified field (text): "Please include your [field name]."
 ** unspecified field (other): "Please indicate [field name]."
 ** invalid text field entries: "[field value] is an invalid [field name]!"
 ** help with format: "Use this format: "
***************************************************************/
var beginRequestAlertForText = "Please include your ";
var beginRequestAlertGeneric = "Please indicate ";
var endRequestAlert = ".";
var beginInvalidAlert = " is an invalid ";
var endInvalidAlert = "!";
var beginFormatAlert = "  Use this format: ";

/***************************************************************
** functions to validate the string or form object passed in,
** and return true or false based on user input
***************************************************************/
function isText(str) {
    return (str != "");
}

function isSelect(formObj) {
    return (formObj.selectedIndex != 0);
}

function isRadio(formObj) {
    for (j=0; j<formObj.length; j++) {
        if (formObj[j].checked) {
            return true;
        }
    }
    return false;
}

function isCheck(formObj, form, begin, num) {
    for (j=begin; j<begin+num; j++) {
        if (form.elements[j].checked) {
            return true;
        }
    }
    return false;
}

function isEmail(str) {
    return ((str != "") && (str.indexOf("@") != -1) && (str.indexOf(".") != -1));
}

function isZipCode(str) {
    var l = str.length;
    if ((l != 5) && (l != 10)) {return false }
    for (j=0; j<l; j++) {
        if ((l == 10) && (j == 5)) {
            if (str.charAt(j) != "-") {return false }
        } else {
            if ((str.charAt(j) < "0") || (str.charAt(j) > "9")) { return false }
        }
    }
    return true;
}

function isPhoneNum(str) {
    if (str.length != 12) { return false }
    for (j=0; j<str.length; j++) {
        if ((j == 3) || (j == 7)) {
            if (str.charAt(j) != "-") { return false }
        } else {
            if ((str.charAt(j) < "0") || (str.charAt(j) > "9")) { return false }
        }
    }
    return true
}

/****************
validate form
****************/
function validateForm(form) {
    var formEltName = "";
    var formObj = "";
    var str = "";
    var realName = "";
    var alertText = "";
    var firstMissingElt = null;
    var hardReturn = "\r\n";
    for (i=0; i<elts.length; i++) {
        formEltName = elts[i].formEltName;
        formObj = eval("form." + formEltName);
        realName = elts[i].realName;

        if (elts[i].eltType == "text") {
            str = formObj.value;
            if (eval(elts[i].upToSnuff)) continue;

            if (str == "") {
                if (allAtOnce) {
                    alertText += beginRequestAlertForText + realName + endRequestAlert + hardReturn;
                    if (firstMissingElt == null) {firstMissingElt = formObj};
                } else {
                    alertText = beginRequestAlertForText + realName + endRequestAlert + hardReturn;
                    alert(alertText);
                }
            } else {
                if (allAtOnce) {
                    alertText += str + beginInvalidAlert + realName + endInvalidAlert + hardReturn;
                } else {
                    alertText = str + beginInvalidAlert + realName + endInvalidAlert + hardReturn;
                }
                if (elts[i].format != null) {
                    alertText += beginFormatAlert + elts[i].format + hardReturn;
                }
                if (allAtOnce) {
                    if (firstMissingElt == null) {firstMissingElt = formObj};
                } else {
                    alert(alertText);
                }
            }
        } else {
            if (eval(elts[i].upToSnuff)) continue;
            if (allAtOnce) {
                alertText += beginRequestAlertGeneric + realName + endRequestAlert + hardReturn;
                if (firstMissingElt == null) {firstMissingElt = formObj};
            } else {
                alertText = beginRequestAlertGeneric + realName + endRequestAlert + hardReturn;
                alert(alertText);
            }
        }
        {
            var goToObj = (allAtOnce) ? firstMissingElt : formObj;
            if (goToObj.select) goToObj.select();
            if (goToObj.focus) goToObj.focus();
        }

        if (!allAtOnce) {return false};
    }
    if (allAtOnce) {
        if (alertText != "") {
            alert(alertText);
            return false;
        }
    }

    return true;
}

//</script>
