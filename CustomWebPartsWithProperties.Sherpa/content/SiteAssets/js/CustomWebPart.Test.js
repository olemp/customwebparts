/// <reference path="SharePoint.d.ts" />
/// <reference path="jquery.d.ts" />
var CustomWebPart;
(function (CustomWebPart) {
    var Test;
    (function (Test) {
        /* HelloWorld */
        /* REQUIRED: None */
        function HelloWorldWebPart(webpart) {
            webpart.instance.html("<p>Hello world</p>");
        }
        Test.HelloWorldWebPart = HelloWorldWebPart;
        /* Permissions */
        /* REQUIRED: [List], [ItemsCount] */
        function ListItems(webpart) {
            var properties = webpart.properties[0];
            jQuery.ajax({
                url: String['format']("{0}/_api/web/lists/getByTitle('{1}')/items?$top={2}", _spPageContextInfo.webAbsoluteUrl, properties["List"], properties["ItemsCount"]),
                type: 'get',
                headers: {
                    'accept': 'application/json;odata=nometadata'
                },
                success: function (d) {
                    var stringBuilder = [];
                    stringBuilder.push("<ul>");
                    jQuery.each(d.value, function (id, val) {
                        stringBuilder.push(String['format']("<li>{0}</li>", val.Title));
                    });
                    stringBuilder.push("<ul>");
                    webpart.instance.html(stringBuilder.join(''));
                }
            });
        }
        Test.ListItems = ListItems;
        /* Permissions */
        /* REQUIRED: None */
        function Permissions(webpart) {
            var properties = webpart.properties[0];
            jQuery.ajax({
                url: String['format']("{0}/_api/web/RoleAssignments?$expand=Member,Member/Users,RoleDefinitionBindings", _spPageContextInfo.webAbsoluteUrl),
                type: 'get',
                headers: {
                    'accept': 'application/json;odata=verbose'
                },
                success: function (d) {
                    var stringBuilder = [];
                    stringBuilder.push("<ul>");
                    jQuery.each(d.d.results, function (id, val) {
                        stringBuilder.push(String['format']("<li>'{0}' has permission '{1}' on this web.</li>", val.Member.Title, val.RoleDefinitionBindings.results[0].Name));
                        if (val.Member.Users && val.Member.Users.results && val.Member.Users.results.length > 0) {
                            jQuery.each(val.Member.Users.results, function (id, val) {
                                var externalUser = (val.LoginName.indexOf("#ext#") != -1) ? " <b>and is an external user</b>" : "";
                                stringBuilder.push(String['format']("<li style='margin-left: 40px;'><i>{0} is a member of this group{1}.</i></li>", val.Title, externalUser));
                            });
                        }
                    });
                    stringBuilder.push("<ul>");
                    webpart.instance.html(stringBuilder.join(''));
                }
            });
        }
        Test.Permissions = Permissions;
        /* PermissionsWithCSS */
        /* REQUIRED: [CSS] */
        function PermissionsWithCSS(webpart) {
            var properties = webpart.properties[0];
            jQuery("head").append(String['format']("<link rel='stylesheet' type='text/css' href='{0}' />", webpart.properties[0]["CSS"]));
            jQuery.ajax({
                url: String['format']("{0}/_api/web/RoleAssignments?$expand=Member,Member/Users,RoleDefinitionBindings", _spPageContextInfo.webAbsoluteUrl),
                type: 'get',
                headers: {
                    'accept': 'application/json;odata=verbose'
                },
                success: function (d) {
                    var stringBuilder = [];
                    stringBuilder.push("<ul id='permissions-web'>");
                    var extUsers = 0;
                    jQuery.each(d.d.results, function (id, val) {
                        stringBuilder.push(String['format']("<li><b>{0}</b> has permission <b>{1}</b> on this web.</li>", val.Member.Title, val.RoleDefinitionBindings.results[0].Name));
                        if (val.Member.Users && val.Member.Users.results && val.Member.Users.results.length > 0) {
                            stringBuilder.push("<li style='margin-left: 40px;'><b>Members:</b></li>");
                            jQuery.each(val.Member.Users.results, function (id, val) {
                                var externalUser = (val.LoginName.indexOf("#ext#") != -1) ? " <b>(external user)</b>" : "";
                                stringBuilder.push(String['format']("<li style='margin-left: 40px;'><i>{0}{1}.</i></li>", val.Title, externalUser));
                                if (externalUser != "") {
                                    extUsers++;
                                }
                            });
                        }
                    });
                    stringBuilder.push("</ul>");
                    stringBuilder.push(String['format']("<p><b>A total of {0} external users has access to this site.</b></p>", extUsers));
                    webpart.instance.html(stringBuilder.join(''));
                }
            });
        }
        Test.PermissionsWithCSS = PermissionsWithCSS;
        /* CalendarItems */
        /* REQUIRED: [Web], [List], [ItemsCount], [Category] */
        function CalendarItems(webpart) {
            var properties = webpart.properties[0];
            jQuery.ajax({
                url: String['format']("{0}{1}/_api/web/lists/getByTitle('{2}')/items?$top={3}&$filter=Category eq '{4}'", _spPageContextInfo.webAbsoluteUrl, properties["Web"], properties["List"], properties["ItemsCount"], properties["Category"]),
                type: 'get',
                headers: {
                    'accept': 'application/json;odata=nometadata'
                },
                success: function (d) {
                    var stringBuilder = [];
                    stringBuilder.push("<ul>");
                    jQuery.each(d.value, function (id, val) {
                        stringBuilder.push(String['format']("<li>{0}</li>", val.Title));
                    });
                    stringBuilder.push("<ul>");
                    webpart.instance.html(stringBuilder.join(''));
                }
            });
        }
        Test.CalendarItems = CalendarItems;
        /* RSSFeed */
        function RSSFeed(webpart) {
            var properties = webpart.properties[0];
            jQuery['getFeed']({
                url: properties["Url"],
                success: function (feed) {
                    console.log(feed.title);
                },
                error: function () {
                    console.log(arguments);
                }
            });
        }
        Test.RSSFeed = RSSFeed;
    })(Test = CustomWebPart.Test || (CustomWebPart.Test = {}));
})(CustomWebPart || (CustomWebPart = {}));
