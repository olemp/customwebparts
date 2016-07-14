/// <reference path="..\typings\main.d.ts" />

namespace OM.CustomWebParts {
    const _config = {
        wp_class: ".custom-webpart",
        linkedit_selector: ".ms-rte-embedcode-linkedit",
        hiddeninput_selector: ".aspNetHidden input",
        toolpane_selector: ".ms-TPBody"
    }

    /**
     * Utilities
     */
    namespace Util {
        /**
         * Retrieves all WebPart definitions from the DOM
         */
        export function GetWebPartsDefinitions(): JQuery {
            return jQuery(_config.wp_class);
        }

        /**
         * Replaces all occurences of a string with another
         * 
         * @param str The string
         * @param f The char/string to replace
         * @param r The char/string to replace it with
         */
        export function ReplaceAll(str: string, f: string, r: string): string {
            return str ? str.replace(new RegExp(f, '\g'), r) : "";
        }

        /**
         * Gets the toolpane for a Webpart
         * 
         * @param webpartid The ID of the webpart
         */
        export function GetToolPaneForWebPart(webpartid: string): JQuery {
            let _wpId = ReplaceAll(webpartid, '-', '_');
            return jQuery(`${_config.toolpane_selector}[id*='${_wpId}']`).first();
        }

        /**
         * Gets the hidden input field for a Webpart
         * 
         * @param webpartid The ID of the webpart
         */
        export function GetHiddenInputFieldForWebPart(webpartid: string): JQuery {
            return jQuery(`${_config.hiddeninput_selector}[name*='${webpartid}']`);
        }

        /**
         * Retrieves selected options from array
         * 
         * @param options Collection of options
         * @param defaultValue The default value
         */
        function GetSelectOptionsFromArray(options: Array<string>, defaultValue: string): string {
            return options.map(o => String.format("<option{0}>{1}</option>", (o == defaultValue) ? " selected" : "", o)).join("");
        }

        /**
         * Retrieves updated webpart html
         * 
         * @param instance The WebPart instance
         */
        function GetUpdatedWebPartHtml(instance: JQuery): string {
            var properties = instance.data("webpart-properties")[0];
            Object.keys(properties).forEach(key => {
                var $input = jQuery(`input.UserInput[name*='EditorZone'][name*='${key}'], select.UserSelect[name*='EditorZone'][name*='${key}']`), eleType = $input.prop("tagName");
                switch (eleType) {
                    case "INPUT": {
                        switch ($input.attr("type")) {
                            case "text": properties[key] = $input.val();
                                break;
                            case "checkbox": properties[key] = $input.prop("checked").toString();
                                break;
                        }
                    }
                        break;
                    case "SELECT": properties[key] = $input.val();;
                        break;
                }
            });
            instance.attr("data-webpart-properties", `[${JSON.stringify(properties)}]`);
            return jQuery('<div>').append(instance.clone()).html();
        }

        /**
         * Logs and an log message
         * 
         * @param msg Message to log
         */
        export function Log(msg: string): void {
            if (window.hasOwnProperty("console") && window.console.info) {
                console.info(msg);
            }
        }

        /**
         * Logs and an error message
         * 
         * @param msg Message to log
         */
        export function Error(msg: string): void {
            if (window.hasOwnProperty("console") && window.console.error) {
                console.error(msg);
            }
        }

        /**
         * Checks if page is in edit mode
         */
        export function InEditMode(): boolean {
            var formName = (typeof window['MSOWebPartPageFormName'] === "string") ? window['MSOWebPartPageFormName'] : "aspnetForm", form = window.document.forms[formName];
            if (form && ((form['MSOLayout_InDesignMode'] && form['MSOLayout_InDesignMode'].value) || (typeof window['MSOLayout_IsWikiEditMode'] === "function" && window['MSOLayout_IsWikiEditMode']()))) {
                return true;
            } else {
                return false;
            }
        }

        /**
         * Renders webpart properties
         * 
         * @param webpart The WebPart to render properties for
         */
        export function RenderWebPartProperties(webpart: WebPart): void {
            var properties = webpart.properties[0], $toolPane = GetToolPaneForWebPart(webpart.id[1]);
            if (Object.keys(properties).length > 0 && $toolPane.length > 0) {
                jQuery(_config.linkedit_selector).hide();
                var props = [];
                for (var i = 0; i < Object.keys(properties).length; i++) {
                    var key = Object.keys(properties)[i], value = properties[key];
                    if (webpart.instance.data("webpart-choices") != null && webpart.instance.data("webpart-choices")[key] != null) {
                        var options = GetSelectOptionsFromArray(webpart.instance.data("webpart-choices")[key].split(","), value);
                        props.push(String.format(Templates.Field_Choice, Util.ReplaceAll(webpart.id[1], '-', '_'), key, options));
                    } else {
                        if (value == "true" || value == "false") {
                            props.push(String.format(Templates.Field_Boolean, Util.ReplaceAll(webpart.id[1], '-', '_'), key, value == "true" ? "checked" : ""))
                        } else {
                            props.push(String.format(Templates.Field_String, Util.ReplaceAll(webpart.id[1], '-', '_'), key, value))
                        }
                    }
                }
                $toolPane.append(String.format(Templates.Container, Util.ReplaceAll(webpart.id[1], '-', '_'), props.join('')));
                var $submit = jQuery("input[type='submit'][name*='OKBtn'], input[type='submit'][name*='AppBtn']");
                $submit.click(() => {
                    GetHiddenInputFieldForWebPart(webpart.id[1]).val(GetUpdatedWebPartHtml(webpart.instance));
                });
            }
        }
    }

    /**
     * Defines a WebPart
     */
    export class WebPart {
        public instance: any;
        public id: Array<string>;
        public renderfunction: string;
        public renderevent: string;
        public properties: Array<Object>;

        /**
         * Renders the webpart
         */
        public render(): void {
            Manager.Render(this);
        }

        /**
         * Constructor
         * 
         * @param elemenent The WebPart DOM Element
         */
        constructor(element: any) {
            this.instance = element;
            this.id = [
                this.instance.parents("div[webpartid]").first().attr("webpartid"),
                this.instance.parents("div[webpartid2]").first().length > 0 ? this.instance.parents("div[webpartid2]").first().attr("webpartid2") : this.instance.parents("div[webpartid]").first().attr("webpartid")
            ]
            this.renderfunction = this.instance.data("webpart-renderfunction");
            this.renderevent = this.instance.data("webpart-renderevent");
            this.properties = this.instance.data("webpart-properties");
        }
    }

    /**
     * A collection of WebPart
     */
    export var WebParts: WebPart[] = [];

    /**
     * Templates for tool pane
     */
    var Templates = {
        Container: "<div> <table cellspacing=\"0\" cellpadding=\"0\" style=\"width:100%;border-collapse:collapse;\"> <tbody> <tr> <td><div class=\"UserSectionTitle\"><a id=\"ctl00_MSOTlPn_EditorZone_Edit0g_{0}_CustomCategory_IMAGEANCHOR\" href=\"#\" onkeydown=\"WebPartMenuKeyboardClick(this, 13, 32, event);\" style=\"cursor:hand\" onclick=\"javascript:MSOTlPn_ToggleDisplay('ctl00_MSOTlPn_EditorZone_Edit0g_{0}_CustomCategory', 'ctl00_MSOTlPn_EditorZone_Edit0g_{0}_CustomCategory_IMAGE', 'ctl00_MSOTlPn_EditorZone_Edit0g_{0}_CustomCategory_ANCHOR', 'Expand category: Custom', 'Collapse category: Custom','ctl00_MSOTlPn_EditorZone_Edit0g_{0}_CustomCategory_IMAGEANCHOR'); return false;\" title=\"Expand category: Custom\">&nbsp;<img id=\"ctl00_MSOTlPn_EditorZone_Edit0g_{0}_CustomCategory_IMAGE\" alt=\"Expand category: Custom\" border=\"0\" src=\"/_layouts/15/images/TPMax2.gif\">&nbsp;</a><a tabindex=\"-1\" onkeydown=\"WebPartMenuKeyboardClick(this, 13, 32, event);\" id=\"ctl00_MSOTlPn_EditorZone_Edit0g_{0}_CustomCategory_ANCHOR\" style=\"cursor:hand\" onclick=\"javascript:MSOTlPn_ToggleDisplay('ctl00_MSOTlPn_EditorZone_Edit0g_{0}_CustomCategory', 'ctl00_MSOTlPn_EditorZone_Edit0g_{0}_CustomCategory_IMAGE', 'ctl00_MSOTlPn_EditorZone_Edit0g_{0}_CustomCategory_ANCHOR', 'Expand category: Custom', 'Collapse category: Custom','ctl00_MSOTlPn_EditorZone_Edit0g_{0}_CustomCategory_IMAGEANCHOR'); return false;\" title=\"Expand category: Custom\"> &nbsp;Custom</a></div></td> </tr> </tbody> </table><div class=\"ms-propGridTable\" id=\"ctl00_MSOTlPn_EditorZone_Edit0g_{0}_CustomCategory\" style=\"display:none;\"> <table cellspacing=\"0\" style=\"border-width:0px;width:100%;border-collapse:collapse;\"> <tbody>{1}</tbody> </table> </div> </div>",
        Field_String: "<tr><td><input type=\"hidden\" name=\"ctl00$MSOTlPn_EditorZone$Edit0g_{0}$ctl11${1}_ROWSTATE\" id=\"ctl00_MSOTlPn_EditorZone_Edit0g_{0}_ctl11_{1}_ROWSTATE\" value=\"0\"><div class=\"UserSectionHead\"><label for=\"ctl00_MSOTlPn_EditorZone_Edit0g_{0}_ctl11_{1}_EDITOR\" title=\"\">{1}</label></div><div class=\"UserSectionBody\"><div class=\"UserControlGroup\"><nobr><input name=\"ctl00$MSOTlPn_EditorZone$Edit0g_{0}$ctl11${1}_EDITOR\" type=\"text\" id=\"ctl00_MSOTlPn_EditorZone_Edit0g_{0}_ctl11_{1}_EDITOR\" class=\"UserInput\" ms-tlpnwiden=\"true\" style=\"width:176px;{1}:ltr;\" value=\"{2}\"></nobr></div></div><div style=\"width:100%\" class=\"UserDottedLine\"></div></td></tr>",
        "Field_Boolean": "<tr><td><input type=\"hidden\" name=\"ctl00$MSOTlPn_EditorZone$Edit0g_{0}$ctl11${1}_ROWSTATE\" id=\"ctl00_MSOTlPn_EditorZone_Edit0g_{0}_ctl11_{1}_ROWSTATE\" value=\"0\"><div class=\"UserSectionHead\"><span onfocus=\"MSOPGrid_HidePrevBuilder()\"><input id=\"ctl00_MSOTlPn_EditorZone_Edit0g_{0}_ctl11_{1}_EDITOR\" type=\"checkbox\" name=\"ctl00$MSOTlPn_EditorZone$Edit0g_{0}$ctl11${1}_EDITOR\" class=\"UserInput\" {2} onclick=\"MSOPGrid_HidePrevBuilder();\"></span>&nbsp;&nbsp;<label for=\"ctl00_MSOTlPn_EditorZone_Edit0g_{0}_ctl11_{1}_EDITOR\" title=\"\">{1}</label></div><div style=\"width:100%\" class=\"UserDottedLine\"></div></td></tr>",
        Field_Choice: "<tr><td><input type=\"hidden\" name=\"ctl00$MSOTlPn_EditorZone$Edit0g_{0}$ctl07${1}_ROWSTATE\" id=\"ctl00_MSOTlPn_EditorZone_Edit0g_{0}_ctl07_{1}_ROWSTATE\" value=\"0\"><div class=\"UserSectionHead\"><label>{1}</label></div><div class=\"UserSectionBody\"><div class=\"UserControlGroup\"><nobr><select name=\"ctl00$MSOTlPn_EditorZone$Edit0g_{0}$ctl07${1}_EDITOR\" id=\"ctl00_MSOTlPn_EditorZone_Edit0g_{0}_ctl07_{1}_EDITOR\" class=\"UserSelect\" onclick=\"MSOPGrid_HidePrevBuilder()\" onfocus=\"MSOPGrid_HidePrevBuilder()\">{2}</select></nobr></div></div><div style=\"width:100%\" class=\"UserDottedLine\"></div></td></tr>"
    };

    /**
     * WebPart Manager, handles rendering
     */
    export namespace Manager {
        export function Init(): void {
            Util.GetWebPartsDefinitions().each(function () {
                try {
                    WebParts.push(new WebPart(jQuery(this)));
                } catch (e) {
                    Util.Error("Error parsing webpart.");
                }
            });
            RenderAllWebParts();
        }
        function RenderAllWebParts(): void {
            WebParts.forEach(wp => {
                if (!wp.renderevent) {
                    wp.render();
                } {
                    ExecuteOrDelayUntilEventNotified(wp.render, wp.renderevent);
                }
            });
        }
        export function Render(webpart: WebPart): void {
            if (!Util.InEditMode()) {
                try {
                    eval(`${webpart.renderfunction}(webpart)`);
                } catch (e) {
                    Util.Error("The render function for one of the webparts doesn't exist, or has a syntax error.");
                }
            } else {
                Util.RenderWebPartProperties(webpart);
            }
        }
    }
}
ExecuteOrDelayUntilBodyLoaded(() => {
    if (!window["_v_dictSod"]["jquery"]) {
        console.error("You need to have a SOD registered for jQuery, and ensure it's loaded.");
        return;
    }
    ExecuteOrDelayUntilScriptLoaded(OM.CustomWebParts.Manager.Init, "jquery");
});
