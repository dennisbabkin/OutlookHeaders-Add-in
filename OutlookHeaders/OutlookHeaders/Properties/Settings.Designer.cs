﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:2.0.50727.9148
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace OutlookHeaders.Properties {
    
    
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "9.0.0.0")]
    internal sealed partial class Settings : global::System.Configuration.ApplicationSettingsBase {
        
        private static Settings defaultInstance = ((Settings)(global::System.Configuration.ApplicationSettingsBase.Synchronized(new Settings())));
        
        public static Settings Default {
            get {
                return defaultInstance;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("True")]
        public bool OutlookHdrsEnabled {
            get {
                return ((bool)(this["OutlookHdrsEnabled"]));
            }
            set {
                this["OutlookHdrsEnabled"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("")]
        public string OutlookHdrsMailHdrs {
            get {
                return ((string)(this["OutlookHdrsMailHdrs"]));
            }
            set {
                this["OutlookHdrsMailHdrs"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("0, 0")]
        public global::System.Drawing.Point OutlookHdrsWndLocation {
            get {
                return ((global::System.Drawing.Point)(this["OutlookHdrsWndLocation"]));
            }
            set {
                this["OutlookHdrsWndLocation"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("0, 0")]
        public global::System.Drawing.Size OutlookHdrsWndSize {
            get {
                return ((global::System.Drawing.Size)(this["OutlookHdrsWndSize"]));
            }
            set {
                this["OutlookHdrsWndSize"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("False")]
        public bool OutlookHdrsWndMax {
            get {
                return ((bool)(this["OutlookHdrsWndMax"]));
            }
            set {
                this["OutlookHdrsWndMax"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        public global::System.Collections.Specialized.StringCollection OutlookHdrsWndLVColSzs {
            get {
                return ((global::System.Collections.Specialized.StringCollection)(this["OutlookHdrsWndLVColSzs"]));
            }
            set {
                this["OutlookHdrsWndLVColSzs"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("False")]
        public bool OutlookHdrsLogWhenSending {
            get {
                return ((bool)(this["OutlookHdrsLogWhenSending"]));
            }
            set {
                this["OutlookHdrsLogWhenSending"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("True")]
        public bool OutlookHdrsUpgrade {
            get {
                return ((bool)(this["OutlookHdrsUpgrade"]));
            }
            set {
                this["OutlookHdrsUpgrade"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("")]
        public string OutlookHdrsVer {
            get {
                return ((string)(this["OutlookHdrsVer"]));
            }
            set {
                this["OutlookHdrsVer"] = value;
            }
        }
    }
}
