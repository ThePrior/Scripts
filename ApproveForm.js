
//See - https://nauzadk.wordpress.com/2012/10/02/working-with-commandaction-and-urlaction/

//This will run on every page in SharePoint - NOT now since called from Calculated Column 
//(MagicLink column) =DppFormBase_FormStatus&"<img src=""/_layouts/images/blank.gif"" onload=""{"&"initApproveForm();}"">"
//Update pdf href to run javascript to alter the forms status to pending approval
//TODO: maybe better to use a button. This href has some weird behaviour.
function initApproveForm() {
    //alert(_spPageContextInfo.webServerRelativeUrl);
    var pdfPath = _spPageContextInfo.webServerRelativeUrl + "PDFs"

    //$("ms-cellstyle ms-vb2").find(); //TODO: restrict the query.

    //Flag to ensure only called once per page load.
    if (typeof initApproveForm.counter === 'undefined') {
        // It has not... perform the initialization
        initApproveForm.counter = 0;

        console.log("in initApproveForm: updating hrefs - " + initApproveForm.counter++);

        $("a[href$='.pdf']").attr("target", "_blank");
        $("a[href$='.pdf']").click(function () {
            viewFormPDF($(this).attr("href"));
            return true;
        });

        $("img[id^='RejectForm_']").click(function () {
            rejectFormNew($(this).data("fhuid"));
            return true;
        });

        $("img[id^='ApproveForm_']").click(function () {
            approveFormNew($(this).data("fhuid"));
            return true;
        });

        $("img[id^='FlagForm_']").click(function () {
            flagFormForReviewNew($(this).data("fhuid"));
            return true;
        });

        // Replace href: /emr/SITE/Forms/223822.txt with link to PDFs
        // CRQ-101749: New href will end with WRSPatientID_PatientLastName_PatientFirstName_PatientDOB_FHUid (fields from EMT_Template content type)
        // CRQ-101749: e.g. 555_Ramsbottom_Douglas_12241983_358322.pdf 
        //$("a[href$='.txt']").attr("target", "_blank");		
        $("a[href$='.txt']").attr('href', function () { return this.href.replace("Forms", "PDFs").replace(".txt", ".pdf"); });

    }

    //$("a[href$='.txt']").css("background-color", "yellow");
}

function DOES_NOT_WORK_YET_viewForm(href) {
    SP.UI.ModalDialog.showModalDialog({
        url: href,
        dialogReturnValueCallback: function (dialogResult, returnValue) {
            //RefreshPage(dialogResult);
            if (dialogResult == SP.UI.DialogResult.OK) {
                SP.UI.Notify.addNotification("Form Viewed", true);
            }
        }
    });
}

function viewFormTxt(href) {
    var indexSlash = href.lastIndexOf("_");
    var indexDot = href.lastIndexOf(".");
    var fhUid = href.slice(indexSlash + 1, indexDot);

    var listId = _spPageContextInfo.pageListId;

    console.log("viewFormPDF - ListId = " + listId + ", fhUid = " + fhUid);

    var setProperties = { 'DppFormBase_FormStatus': 'Pending Approval' };


    //find matching item in SP list and set it's status to "pending approval"		
    //Search by FHUid.
    //Note: Only change status of form if current status is New.
    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml('<View><Query><Where><And><Eq><FieldRef Name="DppFormBase_FHUid" /><Value Type="Integer">' + fhUid + '</Value></Eq><Eq><FieldRef Name="DppFormBase_FormStatus" /><Value Type="Choice">New</Value></Eq></And></Where></Query><RowLimit>1</RowLimit></View>');


    //PageRefresh is required to give visual feedback to user, by changing the Aproval Status field.
    //But need to reload entire page so that jquery can update the href in the list item again!
    //TODO: There may be a better solution using a Button and a SP modal dialog.
    getListItemsCamlQuery(listId, camlQuery, setProperties)
    .then(updateListItems)
    .then(logSuccessWithPageRefresh, logError);

    return false;
}

function approveForm(itemId, listId, siteUrl) {

    console.log("approveForm - ItemId = " + itemId + ", ListId = " + listId + ", SiteUrl = " + siteUrl);

    var setProperties = { 'DppFormBase_FormStatus': 'Approved' };

    //find matching item in SP list - provided status is NOT new - and set it's status to "Approved"		
    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml('<View><Query><Where><And><Eq><FieldRef Name="ID" /><Value Type="Counter">' + itemId + '</Value></Eq><Neq><FieldRef Name="DppFormBase_FormStatus" /><Value Type="Choice">New</Value></Neq></And></Where></Query></View>');
    getListItemsCamlQuery(listId, camlQuery, setProperties)
    .then(processForm);
}

function flagFormForReview(itemId, listId, siteUrl) {

    console.log("flagFormForReview - ItemId = " + itemId + ", ListId = " + listId + ", SiteUrl = " + siteUrl);

    var setProperties = { 'DppFormBase_FormStatus': 'Review' };

    //find matching item in SP list - provided status is NOT new - and set it's status to "Approved"		
    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml('<View><Query><Where><And><Eq><FieldRef Name="ID" /><Value Type="Counter">' + itemId + '</Value></Eq><Neq><FieldRef Name="DppFormBase_FormStatus" /><Value Type="Choice">New</Value></Neq></And></Where></Query></View>');
    getListItemsCamlQuery(listId, camlQuery, setProperties)
    .then(processForm);
}

function viewFormPDF(href) {
    var indexSlash = href.lastIndexOf("_");
    var indexDot = href.lastIndexOf(".");
    var fhUid = href.slice(indexSlash + 1, indexDot);

    var listId = _spPageContextInfo.pageListId;

    console.log("viewFormPDF - ListId = " + listId + ", fhUid = " + fhUid);

    var setProperties = { 'DppFormBase_FormStatus': 'Pending Approval' };


    //find matching item in SP list and set it's status to "pending approval"		
    //Search by FHUid.
    //Note: Only change status of form if current status is New.
    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml('<View><Query><Where><And><Eq><FieldRef Name="DppFormBase_FHUid" /><Value Type="Integer">' + fhUid + '</Value></Eq><Eq><FieldRef Name="DppFormBase_FormStatus" /><Value Type="Choice">New</Value></Eq></And></Where></Query><RowLimit>1</RowLimit></View>');


    //PageRefresh is required to give visual feedback to user, by changing the Aproval Status field.
    //But need to reload entire page so that jquery can update the href in the list item again!
    //TODO: There may be a better solution using a Button and a SP modal dialog.
    getListItemsCamlQuery(listId, camlQuery, setProperties)
    .then(updateListItems)
    .then(logSuccessWithPageRefresh, logError);

    return false;
}

function rejectFormNew(fhUid) {

    var listId = _spPageContextInfo.pageListId;
    console.log("rejectFormNew - FHUid = " + fhUid + ", ListId = " + listId);

    var setProperties = { 'DppFormBase_FormStatus': 'Rejected' };

    //find matching item in SP list - provided status is NOT new - and set it's status to "Rejected"		
    //TODO: Maybe display dialog to collect a reason from the user?
    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml('<View><Query><Where><And><Eq><FieldRef Name="DppFormBase_FHUid" /><Value Type="Integer">' + fhUid + '</Value></Eq><Neq><FieldRef Name="DppFormBase_FormStatus" /><Value Type="Choice">New</Value></Neq></And></Where></Query></View>');
    getListItemsCamlQuery(listId, camlQuery, setProperties)
    .then(processForm);
}


function approveFormNew(fhUid) {

    var listId = _spPageContextInfo.pageListId;
    console.log("approveFormNew - FHUid = " + fhUid + ", ListId = " + listId);

    var setProperties = { 'DppFormBase_FormStatus': 'Approved' };

    //find matching item in SP list - provided status is NOT new - and set it's status to "Approved"		
    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml('<View><Query><Where><And><Eq><FieldRef Name="DppFormBase_FHUid" /><Value Type="Integer">' + fhUid + '</Value></Eq><Neq><FieldRef Name="DppFormBase_FormStatus" /><Value Type="Choice">New</Value></Neq></And></Where></Query></View>');
    getListItemsCamlQuery(listId, camlQuery, setProperties)
    .then(processForm);
}

function flagFormForReviewNew(fhUid) {

    var listId = _spPageContextInfo.pageListId;
    console.log("flagFormForReviewNew - FHUid = " + fhUid + ", ListId = " + listId);

    var setProperties = { 'DppFormBase_FormStatus': 'Review' };

    //find matching item in SP list - provided status is NOT new - and set it's status to "Review"		
    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml('<View><Query><Where><And><Eq><FieldRef Name="DppFormBase_FHUid" /><Value Type="Integer">' + fhUid + '</Value></Eq><Neq><FieldRef Name="DppFormBase_FormStatus" /><Value Type="Choice">New</Value></Neq></And></Where></Query></View>');
    getListItemsCamlQuery(listId, camlQuery, setProperties)
    .then(processForm);
}

function approveForm(itemId, listId, siteUrl) {

    console.log("approveForm - ItemId = " + itemId + ", ListId = " + listId + ", SiteUrl = " + siteUrl);

    var setProperties = { 'DppFormBase_FormStatus': 'Approved' };

    //find matching item in SP list - provided status is NOT new - and set it's status to "Approved"		
    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml('<View><Query><Where><And><Eq><FieldRef Name="ID" /><Value Type="Counter">' + itemId + '</Value></Eq><Neq><FieldRef Name="DppFormBase_FormStatus" /><Value Type="Choice">New</Value></Neq></And></Where></Query></View>');
    getListItemsCamlQuery(listId, camlQuery, setProperties)
    .then(processForm);
}

function rejectForm(itemId, listId, siteUrl) {

    console.log("rejectForm - ItemId = " + itemId + ", ListId = " + listId + ", SiteUrl = " + siteUrl);

    var setProperties = { 'DppFormBase_FormStatus': 'Rejected' };

    //find matching item in SP list - provided status is NOT new - and set it's status to "Rejected"		
    //TODO: Maybe display dialog to collect a reason from the user?
    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml('<View><Query><Where><And><Eq><FieldRef Name="ID" /><Value Type="Counter">' + itemId + '</Value></Eq><Neq><FieldRef Name="DppFormBase_FormStatus" /><Value Type="Choice">New</Value></Neq></And></Where></Query></View>');
    getListItemsCamlQuery(listId, camlQuery, setProperties)
    .then(processForm);
}

function processForm(items, itemProperties) {

    var deferred = $.Deferred();
    deferred.fail(function (error) {
        SP.UI.Notify.addNotification(error, false);
    });
    deferred.done(logSuccessWithPageRefresh);

    var ctx = items.get_context();

    var itemCount = 0;
    //prepare to update list items 
    items.get_data().forEach(function (item) {
        itemCount++;
        for (var propName in itemProperties) {
            item.set_item(propName, itemProperties[propName])
        }
        item.update();
    });


    if (itemCount == 0) {
        deferred.reject("Form PDF must be viewed before its status can be changed.");
    } else {
        //submit request
        ctx.executeQueryAsync(function () {
            deferred.resolve(itemCount + ' item(s) updated');
        },
		function (sender, args) {
		    deferred.reject(args.get_message());
		});
    }

    return deferred.promise();

}

// Ensure the SP JSOM is loaded before we worry about our stuff
jQuery(document).ready(function () {
    //ExecuteOrDelayUntilScriptLoaded(initApproveForm, "sp.js");
});


