/// <reference path="SharePoint.d.ts" />
/// <reference path="jquery.d.ts" />
var CustomWebPart;
(function (CustomWebPart) {
    var Util;
    (function (Util) {
        function GetWebPartsDefinitions() {
            return jQuery(Properties.WebPartClass);
        }
        Util.GetWebPartsDefinitions = GetWebPartsDefinitions;
        function ReplaceAll(str, f, r) {
            return str ? str.replace(new RegExp(f, '\g'), r) : "";
        }
        Util.ReplaceAll = ReplaceAll;
        function GetToolPaneForWebPart(webpartid) {
            return jQuery(".ms-TPBody[id*='" + ReplaceAll(webpartid, '-', '_') + "']").first();
        }
        Util.GetToolPaneForWebPart = GetToolPaneForWebPart;
        function GetHiddenInputFieldForWebPart(webpartid) {
            return jQuery(".aspNetHidden input[name*='" + webpartid + "']");
        }
        Util.GetHiddenInputFieldForWebPart = GetHiddenInputFieldForWebPart;
        function GetSelectOptionsFromArray(options, defaultValue) {
            var html = "";
            options.forEach(function (val, id) {
                html += String['format']("<option{0}>{1}</option>", (val == defaultValue) ? " selected" : "", val);
            });
            return html;
        }
        function GetUpdatedWebPartHtml(instance) {
            var properties = instance.data("webpart-properties")[0];
            for (var i = 0; i < Object.keys(properties).length; i++) {
                var key = Object.keys(properties)[i];
                var $input = jQuery("input.UserInput[name*='EditorZone'][name*='" + key + "'], select.UserSelect[name*='EditorZone'][name*='" + key + "']");
                var elementType = $input.prop("tagName");
                switch (elementType) {
                    case "INPUT":
                        {
                            switch ($input.attr("type")) {
                                case "text":
                                    properties[key] = $input.val();
                                    break;
                                case "checkbox":
                                    properties[key] = $input.prop("checked").toString();
                                    break;
                            }
                        }
                        break;
                    case "SELECT":
                        properties[key] = $input.val();
                        ;
                        break;
                }
            }
            instance.attr("data-webpart-properties", "[" + JSON.stringify(properties) + "]");
            return $('<div>').append(instance.clone()).html();
        }
        function Log(message) {
            console.info(message);
        }
        Util.Log = Log;
        function Error(message) {
            console.error(message);
        }
        Util.Error = Error;
        function InEditMode() {
            var formName = (typeof window['MSOWebPartPageFormName'] === "string") ? window['MSOWebPartPageFormName'] : "aspnetForm";
            var form = window.document.forms[formName];
            if (form && ((form['MSOLayout_InDesignMode'] && form['MSOLayout_InDesignMode'].value) || (typeof window['MSOLayout_IsWikiEditMode'] === "function" && window['MSOLayout_IsWikiEditMode']()))) {
                return true;
            }
            else {
                return false;
            }
        }
        Util.InEditMode = InEditMode;
        function RenderWebPartProperties(webpart) {
            var properties = webpart.properties[0];
            var $toolPane = GetToolPaneForWebPart(webpart.id[1]);
            if (Object.keys(properties).length > 0 && $toolPane.length > 0) {
                jQuery(".ms-rte-embedcode-linkedit").hide();
                jQuery.getJSON(_spPageContextInfo.siteAbsoluteUrl + Properties.HtmlRootPath + "customproperties.txt", function (d) {
                    var props = [];
                    for (var i = 0; i < Object.keys(properties).length; i++) {
                        var key = Object.keys(properties)[i];
                        var value = properties[key];
                        if (webpart.instance.data("webpart-choices") != null && webpart.instance.data("webpart-choices")[key] != null) {
                            var options = GetSelectOptionsFromArray(webpart.instance.data("webpart-choices")[key].split(","), value);
                            props.push(String['format'](d.Field_Choice, Util.ReplaceAll(webpart.id[1], '-', '_'), key, options));
                        }
                        else {
                            if (value == "true" || value == "false") {
                                props.push(String['format'](d.Field_Boolean, Util.ReplaceAll(webpart.id[1], '-', '_'), key, value == "true" ? "checked" : ""));
                            }
                            else {
                                props.push(String['format'](d.Field_String, Util.ReplaceAll(webpart.id[1], '-', '_'), key, value));
                            }
                        }
                    }
                    $toolPane.append(String['format'](d.Container, Util.ReplaceAll(webpart.id[1], '-', '_'), props.join('')));
                    var $submit = jQuery("input[type='submit'][name*='OKBtn'], input[type='submit'][name*='AppBtn']");
                    $submit.click(function (event, args) {
                        GetHiddenInputFieldForWebPart(webpart.id[1]).val(GetUpdatedWebPartHtml(webpart.instance));
                    });
                });
            }
        }
        Util.RenderWebPartProperties = RenderWebPartProperties;
    })(Util || (Util = {}));
    var Properties;
    (function (Properties) {
        Properties.WebPartClass = '.custom-webpart';
        Properties.HtmlRootPath = "/siteassets/customwebparts/html/";
    })(Properties = CustomWebPart.Properties || (CustomWebPart.Properties = {}));
    var Model;
    (function (Model) {
        var WebPart = (function () {
            function WebPart(element) {
                this.instance = element;
                this.id = [
                    this.instance.parents("div[webpartid]").first().attr("webpartid"),
                    this.instance.parents("div[webpartid2]").first().length > 0 ? this.instance.parents("div[webpartid2]").first().attr("webpartid2") : this.instance.parents("div[webpartid]").first().attr("webpartid")
                ];
                this.renderfunction = this.instance.data("webpart-renderfunction");
                this.properties = this.instance.data("webpart-properties");
            }
            WebPart.prototype.render = function () {
                Manager.Render(this);
            };
            return WebPart;
        })();
        Model.WebPart = WebPart;
        Model.WebParts = [];
    })(Model = CustomWebPart.Model || (CustomWebPart.Model = {}));
    var Manager;
    (function (Manager) {
        function Init() {
            Util.GetWebPartsDefinitions().each(function () {
                try {
                    Model.WebParts.push(new Model.WebPart(jQuery(this)));
                }
                catch (e) {
                    Util.Error("Error parsing webpart.");
                }
            });
            RenderAllWebParts();
        }
        Manager.Init = Init;
        function RenderAllWebParts() {
            for (var i in Model.WebParts) {
                Model.WebParts[i].render();
            }
        }
        function Render(webpart) {
            if (!Util.InEditMode()) {
                try {
                    eval(webpart.renderfunction + "(webpart)");
                }
                catch (e) {
                    Util.Error("The render function for one of the webparts doesn't exist, or has a syntax error.");
                }
            }
            else {
                Util.RenderWebPartProperties(webpart);
            }
        }
        Manager.Render = Render;
    })(Manager = CustomWebPart.Manager || (CustomWebPart.Manager = {}));
})(CustomWebPart || (CustomWebPart = {}));
jQuery(function () {
    CustomWebPart.Manager.Init();
});
