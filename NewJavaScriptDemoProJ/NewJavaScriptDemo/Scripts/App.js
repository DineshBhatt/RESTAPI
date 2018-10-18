'use strict';

ExecuteOrDelayUntilScriptLoaded(initializePage, "sp.js");

function initializePage()
{
    var context = SP.ClientContext.get_current();
    var user = context.get_web().get_currentUser();

    // This code runs when the DOM is ready and creates a context object which is needed to use the SharePoint object model
    $(document).ready(function () {
        getUserName();
        $("#Button1").click(GetList);
        $("#Button2").click(CreateList);
        $("#Button3").click(CreateField);
    });

    // This function prepares, loads, and then executes a SharePoint query to get the current users information
    function getUserName() {
        context.load(user);
        context.executeQueryAsync(onGetUserNameSuccess, onGetUserNameFail);
    }
    function GetList() {
        var call1 = jQuery.ajax({
            url: _spPageContextInfo.webAbsoluteUrl + "/_api/Web/?$select=Title",
            type: "GET",
            dataType: "json",
            headers: {
                Accept: "application/json;odata=verbose"
            }
        });
        var call2 = jQuery.ajax({
            url: _spPageContextInfo.webAbsoluteUrl + "/_api/Web/Lists?$select=Title,Hidden,ItemCount&$orderby=ItemCount&$filter=((Hidden eq false) and (ItemCount gt 0))",
            type: "GET",
            dataType: "json",
            headers: {
                Accept: "application/json;odata=verbose"
            }
        });
        var calls = jQuery.when(call1, call2);
        calls.done(function (callback1, callback2) {
            var message = jQuery("#message");
            message.text("Lists in " + callback1[0].d.Title);
            message.append("<br/>");
            jQuery.each(callback2[0].d.results, function (index, value) {
                message.append(String.format("List {0} has {1} items and is {2} hidden",
                    value.Title, value.ItemCount, value.Hidden ? "" : "not"));
                message.append("<br/>");
            });
        });
        calls.fail(function (jqXHR, textStatus, errorThrown) {
            var response = JSON.parse(jqXHR.responseText);
            var message = response ? response.error.message.value : textStatus;
            alert("Call failed. Error: " + message);
        });
    }
    function CreateList()
    {
        var Call1 = jQuery.ajax({
            url: _spPageContextInfo.webAbsoluteUrl + "/_api/Web/Lists",
            type: "POST",
            data: JSON.stringify({
                "__metadata": { type: "SP.List" },
                BaseTemplate: SP.ListTemplateType.tasks,
                Title: "DashBoard"
            }),
            headers: {
                Accept: "application/json;odata:verbose",
                "Content-Type": "application/json;odata=verbose",
                "X-RequestDigest": jQuery("#__REQUESTDIGEST").val()
            }
            
        });
        Call1.done(function (data, textStatus, jqXHR) {
            var message = jQuery("#message");
            message.append("List Is created");
            message.append("<br/>")
        });
        Call1.fail(function (jqXHR, textStatus, errorThrown) {
            var response = JSON.parse(jqXHR.responseText);
            var message = response ? response.error.message.value : textStatus;
            alert("Call failed. Error: " + message);
        });
    }

    function CreateField() {
        var find = jQuery.ajax({
            url: _spPageContextInfo.webAbsoluteUrl + "/_api/Web/Lists/GetByTitle('DashBoard')/Fields",
            type: "POST",
            data: JSON.stringify({
                "__metadata": { type: "SP.Field" },
                Title: "RecentUpdate",
                FieldTypeKind: 2,
                StaticName : "RecentUpdate"
                
            }),
            headers: {
                Accept: "application/json;odata:verbose",
                "Content-Type": "application/json;odata=verbose",
                "X-RequestDigest": jQuery("#__REQUESTDIGEST").val()
            }
        });
        find.done(function (data, textStatus, jqXHR) {
            var message = jQuery("#message");
            message.append("List Is created");
            message.append("<br/>")
        });
        find.fail(function (jqXHR, textStatus, errorThrown) {
            var response = JSON.parse(jqXHR.responseText);
            var message = response ? response.error.message.value : textStatus;
            alert("Call failed. Error: " + message);
        });
    }
    // This function is executed if the above call is successful
    // It replaces the contents of the 'message' element with the user name
    function onGetUserNameSuccess() {
        $('#message').text('Hello ' + user.get_title());
    }

    // This function is executed if the above call fails
    function onGetUserNameFail(sender, args) {
        alert('Failed to get user name. Error:' + args.get_message());
    }
}
