var OM;
(function (OM) {
    var CustomWebParts;
    (function (CustomWebParts) {
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
                return options.map(function (o) { return String['format']("<option{0}>{1}</option>", (o == defaultValue) ? " selected" : "", o); }).join("");
            }
            function GetUpdatedWebPartHtml(instance) {
                var properties = instance.data("webpart-properties")[0];
                Object.keys(properties).forEach(function (key) {
                    var $input = jQuery("input.UserInput[name*='EditorZone'][name*='" + key + "'], select.UserSelect[name*='EditorZone'][name*='" + key + "']"), elementType = $input.prop("tagName");
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
                });
                instance.attr("data-webpart-properties", "[" + JSON.stringify(properties) + "]");
                return $('<div>').append(instance.clone()).html();
            }
            function Log(message) {
                if (window.hasOwnProperty("console") && window.console.info) {
                    console.info(message);
                }
            }
            Util.Log = Log;
            function Error(message) {
                if (window.hasOwnProperty("console") && window.console.error) {
                    console.error(message);
                }
            }
            Util.Error = Error;
            function InEditMode() {
                var formName = (typeof window['MSOWebPartPageFormName'] === "string") ? window['MSOWebPartPageFormName'] : "aspnetForm", form = window.document.forms[formName];
                if (form && ((form['MSOLayout_InDesignMode'] && form['MSOLayout_InDesignMode'].value) || (typeof window['MSOLayout_IsWikiEditMode'] === "function" && window['MSOLayout_IsWikiEditMode']()))) {
                    return true;
                }
                else {
                    return false;
                }
            }
            Util.InEditMode = InEditMode;
            function RenderWebPartProperties(webpart) {
                var properties = webpart.properties[0], $toolPane = GetToolPaneForWebPart(webpart.id[1]);
                if (Object.keys(properties).length > 0 && $toolPane.length > 0) {
                    jQuery(".ms-rte-embedcode-linkedit").hide();
                    var props = [];
                    for (var i = 0; i < Object.keys(properties).length; i++) {
                        var key = Object.keys(properties)[i], value = properties[key];
                        if (webpart.instance.data("webpart-choices") != null && webpart.instance.data("webpart-choices")[key] != null) {
                            var options = GetSelectOptionsFromArray(webpart.instance.data("webpart-choices")[key].split(","), value);
                            props.push(String['format'](Templates.Field_Choice, Util.ReplaceAll(webpart.id[1], '-', '_'), key, options));
                        }
                        else {
                            if (value == "true" || value == "false") {
                                props.push(String['format'](Templates.Field_Boolean, Util.ReplaceAll(webpart.id[1], '-', '_'), key, value == "true" ? "checked" : ""));
                            }
                            else {
                                props.push(String['format'](Templates.Field_String, Util.ReplaceAll(webpart.id[1], '-', '_'), key, value));
                            }
                        }
                    }
                    $toolPane.append(String['format'](Templates.Container, Util.ReplaceAll(webpart.id[1], '-', '_'), props.join('')));
                    var $submit = jQuery("input[type='submit'][name*='OKBtn'], input[type='submit'][name*='AppBtn']");
                    $submit.click(function (event, args) {
                        GetHiddenInputFieldForWebPart(webpart.id[1]).val(GetUpdatedWebPartHtml(webpart.instance));
                    });
                }
            }
            Util.RenderWebPartProperties = RenderWebPartProperties;
        })(Util || (Util = {}));
        var Properties;
        (function (Properties) {
            Properties.WebPartClass = '.custom-webpart';
        })(Properties || (Properties = {}));
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
        }());
        CustomWebParts.WebPart = WebPart;
        CustomWebParts.WebParts = [];
        var Templates = {
            Container: "<div> <table cellspacing=\"0\" cellpadding=\"0\" style=\"width:100%;border-collapse:collapse;\"> <tbody> <tr> <td><div class=\"UserSectionTitle\"><a id=\"ctl00_MSOTlPn_EditorZone_Edit0g_{0}_CustomCategory_IMAGEANCHOR\" href=\"#\" onkeydown=\"WebPartMenuKeyboardClick(this, 13, 32, event);\" style=\"cursor:hand\" onclick=\"javascript:MSOTlPn_ToggleDisplay('ctl00_MSOTlPn_EditorZone_Edit0g_{0}_CustomCategory', 'ctl00_MSOTlPn_EditorZone_Edit0g_{0}_CustomCategory_IMAGE', 'ctl00_MSOTlPn_EditorZone_Edit0g_{0}_CustomCategory_ANCHOR', 'Expand category: Custom', 'Collapse category: Custom','ctl00_MSOTlPn_EditorZone_Edit0g_{0}_CustomCategory_IMAGEANCHOR'); return false;\" title=\"Expand category: Custom\">&nbsp;<img id=\"ctl00_MSOTlPn_EditorZone_Edit0g_{0}_CustomCategory_IMAGE\" alt=\"Expand category: Custom\" border=\"0\" src=\"/_layouts/15/images/TPMax2.gif\">&nbsp;</a><a tabindex=\"-1\" onkeydown=\"WebPartMenuKeyboardClick(this, 13, 32, event);\" id=\"ctl00_MSOTlPn_EditorZone_Edit0g_{0}_CustomCategory_ANCHOR\" style=\"cursor:hand\" onclick=\"javascript:MSOTlPn_ToggleDisplay('ctl00_MSOTlPn_EditorZone_Edit0g_{0}_CustomCategory', 'ctl00_MSOTlPn_EditorZone_Edit0g_{0}_CustomCategory_IMAGE', 'ctl00_MSOTlPn_EditorZone_Edit0g_{0}_CustomCategory_ANCHOR', 'Expand category: Custom', 'Collapse category: Custom','ctl00_MSOTlPn_EditorZone_Edit0g_{0}_CustomCategory_IMAGEANCHOR'); return false;\" title=\"Expand category: Custom\"> &nbsp;Custom</a></div></td> </tr> </tbody> </table><div class=\"ms-propGridTable\" id=\"ctl00_MSOTlPn_EditorZone_Edit0g_{0}_CustomCategory\" style=\"display:none;\"> <table cellspacing=\"0\" style=\"border-width:0px;width:100%;border-collapse:collapse;\"> <tbody>{1}</tbody> </table> </div> </div>",
            Field_String: "<tr><td><input type=\"hidden\" name=\"ctl00$MSOTlPn_EditorZone$Edit0g_{0}$ctl11${1}_ROWSTATE\" id=\"ctl00_MSOTlPn_EditorZone_Edit0g_{0}_ctl11_{1}_ROWSTATE\" value=\"0\"><div class=\"UserSectionHead\"><label for=\"ctl00_MSOTlPn_EditorZone_Edit0g_{0}_ctl11_{1}_EDITOR\" title=\"\">{1}</label></div><div class=\"UserSectionBody\"><div class=\"UserControlGroup\"><nobr><input name=\"ctl00$MSOTlPn_EditorZone$Edit0g_{0}$ctl11${1}_EDITOR\" type=\"text\" id=\"ctl00_MSOTlPn_EditorZone_Edit0g_{0}_ctl11_{1}_EDITOR\" class=\"UserInput\" ms-tlpnwiden=\"true\" style=\"width:176px;{1}:ltr;\" value=\"{2}\"></nobr></div></div><div style=\"width:100%\" class=\"UserDottedLine\"></div></td></tr>",
            "Field_Boolean": "<tr><td><input type=\"hidden\" name=\"ctl00$MSOTlPn_EditorZone$Edit0g_{0}$ctl11${1}_ROWSTATE\" id=\"ctl00_MSOTlPn_EditorZone_Edit0g_{0}_ctl11_{1}_ROWSTATE\" value=\"0\"><div class=\"UserSectionHead\"><span onfocus=\"MSOPGrid_HidePrevBuilder()\"><input id=\"ctl00_MSOTlPn_EditorZone_Edit0g_{0}_ctl11_{1}_EDITOR\" type=\"checkbox\" name=\"ctl00$MSOTlPn_EditorZone$Edit0g_{0}$ctl11${1}_EDITOR\" class=\"UserInput\" {2} onclick=\"MSOPGrid_HidePrevBuilder();\"></span>&nbsp;&nbsp;<label for=\"ctl00_MSOTlPn_EditorZone_Edit0g_{0}_ctl11_{1}_EDITOR\" title=\"\">{1}</label></div><div style=\"width:100%\" class=\"UserDottedLine\"></div></td></tr>",
            Field_Choice: "<tr><td><input type=\"hidden\" name=\"ctl00$MSOTlPn_EditorZone$Edit0g_{0}$ctl07${1}_ROWSTATE\" id=\"ctl00_MSOTlPn_EditorZone_Edit0g_{0}_ctl07_{1}_ROWSTATE\" value=\"0\"><div class=\"UserSectionHead\"><label>{1}</label></div><div class=\"UserSectionBody\"><div class=\"UserControlGroup\"><nobr><select name=\"ctl00$MSOTlPn_EditorZone$Edit0g_{0}$ctl07${1}_EDITOR\" id=\"ctl00_MSOTlPn_EditorZone_Edit0g_{0}_ctl07_{1}_EDITOR\" class=\"UserSelect\" onclick=\"MSOPGrid_HidePrevBuilder()\" onfocus=\"MSOPGrid_HidePrevBuilder()\">{2}</select></nobr></div></div><div style=\"width:100%\" class=\"UserDottedLine\"></div></td></tr>"
        };
        var Manager;
        (function (Manager) {
            function Init() {
                Util.GetWebPartsDefinitions().each(function () {
                    try {
                        CustomWebParts.WebParts.push(new WebPart(jQuery(this)));
                    }
                    catch (e) {
                        Util.Error("Error parsing webpart.");
                    }
                });
                RenderAllWebParts();
            }
            Manager.Init = Init;
            function RenderAllWebParts() {
                CustomWebParts.WebParts.forEach(function (wp) { return wp.Render(); });
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
        })(Manager = CustomWebParts.Manager || (CustomWebParts.Manager = {}));
    })(CustomWebParts = OM.CustomWebParts || (OM.CustomWebParts = {}));
})(OM || (OM = {}));
ExecuteOrDelayUntilBodyLoaded(function () {
    if (!window["_v_dictSod"]["jquery"]) {
        console.error("You need to have a SOD registered for jQuery, and ensure it's loaded.");
        return;
    }
    ExecuteOrDelayUntilScriptLoaded(OM.CustomWebParts.Manager.Init, "jquery");
});
