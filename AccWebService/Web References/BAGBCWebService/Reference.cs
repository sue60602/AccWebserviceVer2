﻿//------------------------------------------------------------------------------
// <auto-generated>
//     這段程式碼是由工具產生的。
//     執行階段版本:4.0.30319.42000
//
//     對這個檔案所做的變更可能會造成錯誤的行為，而且如果重新產生程式碼，
//     變更將會遺失。
// </auto-generated>
//------------------------------------------------------------------------------

// 
// 原始程式碼已由 Microsoft.VSDesigner 自動產生，版本 4.0.30319.42000。
// 
#pragma warning disable 1591

namespace AccWebService.BAGBCWebService {
    using System;
    using System.Web.Services;
    using System.Diagnostics;
    using System.Web.Services.Protocols;
    using System.Xml.Serialization;
    using System.ComponentModel;
    
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.7.2046.0")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Web.Services.WebServiceBindingAttribute(Name="GBCWebServiceSoap", Namespace="http://tempuri.org/")]
    public partial class GBCWebService : System.Web.Services.Protocols.SoapHttpClientProtocol {
        
        private System.Threading.SendOrPostCallback GetVw_GBCVisaDetailJSONOperationCompleted;
        
        private System.Threading.SendOrPostCallback FillVouNoOperationCompleted;
        
        private System.Threading.SendOrPostCallback GetYearOperationCompleted;
        
        private System.Threading.SendOrPostCallback GetAcmWordNumOperationCompleted;
        
        private System.Threading.SendOrPostCallback GetAccKindOperationCompleted;
        
        private System.Threading.SendOrPostCallback GetAccCountOperationCompleted;
        
        private System.Threading.SendOrPostCallback GetAccDetailOperationCompleted;
        
        private System.Threading.SendOrPostCallback GetByPrimaryKeyOperationCompleted;
        
        private System.Threading.SendOrPostCallback GetByKindOperationCompleted;
        
        private bool useDefaultCredentialsSetExplicitly;
        
        /// <remarks/>
        public GBCWebService() {
            this.Url = global::AccWebService.Properties.Settings.Default.AccWebService_BAGBCWebService_GBCWebService;
            if ((this.IsLocalFileSystemWebService(this.Url) == true)) {
                this.UseDefaultCredentials = true;
                this.useDefaultCredentialsSetExplicitly = false;
            }
            else {
                this.useDefaultCredentialsSetExplicitly = true;
            }
        }
        
        public new string Url {
            get {
                return base.Url;
            }
            set {
                if ((((this.IsLocalFileSystemWebService(base.Url) == true) 
                            && (this.useDefaultCredentialsSetExplicitly == false)) 
                            && (this.IsLocalFileSystemWebService(value) == false))) {
                    base.UseDefaultCredentials = false;
                }
                base.Url = value;
            }
        }
        
        public new bool UseDefaultCredentials {
            get {
                return base.UseDefaultCredentials;
            }
            set {
                base.UseDefaultCredentials = value;
                this.useDefaultCredentialsSetExplicitly = true;
            }
        }
        
        /// <remarks/>
        public event GetVw_GBCVisaDetailJSONCompletedEventHandler GetVw_GBCVisaDetailJSONCompleted;
        
        /// <remarks/>
        public event FillVouNoCompletedEventHandler FillVouNoCompleted;
        
        /// <remarks/>
        public event GetYearCompletedEventHandler GetYearCompleted;
        
        /// <remarks/>
        public event GetAcmWordNumCompletedEventHandler GetAcmWordNumCompleted;
        
        /// <remarks/>
        public event GetAccKindCompletedEventHandler GetAccKindCompleted;
        
        /// <remarks/>
        public event GetAccCountCompletedEventHandler GetAccCountCompleted;
        
        /// <remarks/>
        public event GetAccDetailCompletedEventHandler GetAccDetailCompleted;
        
        /// <remarks/>
        public event GetByPrimaryKeyCompletedEventHandler GetByPrimaryKeyCompleted;
        
        /// <remarks/>
        public event GetByKindCompletedEventHandler GetByKindCompleted;
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/GetVw_GBCVisaDetailJSON", RequestNamespace="http://tempuri.org/", ResponseNamespace="http://tempuri.org/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string GetVw_GBCVisaDetailJSON(string acmWordNum) {
            object[] results = this.Invoke("GetVw_GBCVisaDetailJSON", new object[] {
                        acmWordNum});
            return ((string)(results[0]));
        }
        
        /// <remarks/>
        public void GetVw_GBCVisaDetailJSONAsync(string acmWordNum) {
            this.GetVw_GBCVisaDetailJSONAsync(acmWordNum, null);
        }
        
        /// <remarks/>
        public void GetVw_GBCVisaDetailJSONAsync(string acmWordNum, object userState) {
            if ((this.GetVw_GBCVisaDetailJSONOperationCompleted == null)) {
                this.GetVw_GBCVisaDetailJSONOperationCompleted = new System.Threading.SendOrPostCallback(this.OnGetVw_GBCVisaDetailJSONOperationCompleted);
            }
            this.InvokeAsync("GetVw_GBCVisaDetailJSON", new object[] {
                        acmWordNum}, this.GetVw_GBCVisaDetailJSONOperationCompleted, userState);
        }
        
        private void OnGetVw_GBCVisaDetailJSONOperationCompleted(object arg) {
            if ((this.GetVw_GBCVisaDetailJSONCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.GetVw_GBCVisaDetailJSONCompleted(this, new GetVw_GBCVisaDetailJSONCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/FillVouNo", RequestNamespace="http://tempuri.org/", ResponseNamespace="http://tempuri.org/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string FillVouNo(string accYear, string acmWordNum, string accKind, string accCount, string accDetail, string vouNo, string vouDate, string passNo, string passDate) {
            object[] results = this.Invoke("FillVouNo", new object[] {
                        accYear,
                        acmWordNum,
                        accKind,
                        accCount,
                        accDetail,
                        vouNo,
                        vouDate,
                        passNo,
                        passDate});
            return ((string)(results[0]));
        }
        
        /// <remarks/>
        public void FillVouNoAsync(string accYear, string acmWordNum, string accKind, string accCount, string accDetail, string vouNo, string vouDate, string passNo, string passDate) {
            this.FillVouNoAsync(accYear, acmWordNum, accKind, accCount, accDetail, vouNo, vouDate, passNo, passDate, null);
        }
        
        /// <remarks/>
        public void FillVouNoAsync(string accYear, string acmWordNum, string accKind, string accCount, string accDetail, string vouNo, string vouDate, string passNo, string passDate, object userState) {
            if ((this.FillVouNoOperationCompleted == null)) {
                this.FillVouNoOperationCompleted = new System.Threading.SendOrPostCallback(this.OnFillVouNoOperationCompleted);
            }
            this.InvokeAsync("FillVouNo", new object[] {
                        accYear,
                        acmWordNum,
                        accKind,
                        accCount,
                        accDetail,
                        vouNo,
                        vouDate,
                        passNo,
                        passDate}, this.FillVouNoOperationCompleted, userState);
        }
        
        private void OnFillVouNoOperationCompleted(object arg) {
            if ((this.FillVouNoCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.FillVouNoCompleted(this, new FillVouNoCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/GetYear", RequestNamespace="http://tempuri.org/", ResponseNamespace="http://tempuri.org/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string[] GetYear() {
            object[] results = this.Invoke("GetYear", new object[0]);
            return ((string[])(results[0]));
        }
        
        /// <remarks/>
        public void GetYearAsync() {
            this.GetYearAsync(null);
        }
        
        /// <remarks/>
        public void GetYearAsync(object userState) {
            if ((this.GetYearOperationCompleted == null)) {
                this.GetYearOperationCompleted = new System.Threading.SendOrPostCallback(this.OnGetYearOperationCompleted);
            }
            this.InvokeAsync("GetYear", new object[0], this.GetYearOperationCompleted, userState);
        }
        
        private void OnGetYearOperationCompleted(object arg) {
            if ((this.GetYearCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.GetYearCompleted(this, new GetYearCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/GetAcmWordNum", RequestNamespace="http://tempuri.org/", ResponseNamespace="http://tempuri.org/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string[] GetAcmWordNum(string accYear) {
            object[] results = this.Invoke("GetAcmWordNum", new object[] {
                        accYear});
            return ((string[])(results[0]));
        }
        
        /// <remarks/>
        public void GetAcmWordNumAsync(string accYear) {
            this.GetAcmWordNumAsync(accYear, null);
        }
        
        /// <remarks/>
        public void GetAcmWordNumAsync(string accYear, object userState) {
            if ((this.GetAcmWordNumOperationCompleted == null)) {
                this.GetAcmWordNumOperationCompleted = new System.Threading.SendOrPostCallback(this.OnGetAcmWordNumOperationCompleted);
            }
            this.InvokeAsync("GetAcmWordNum", new object[] {
                        accYear}, this.GetAcmWordNumOperationCompleted, userState);
        }
        
        private void OnGetAcmWordNumOperationCompleted(object arg) {
            if ((this.GetAcmWordNumCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.GetAcmWordNumCompleted(this, new GetAcmWordNumCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/GetAccKind", RequestNamespace="http://tempuri.org/", ResponseNamespace="http://tempuri.org/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string[] GetAccKind(string accYear, string acmWordNum) {
            object[] results = this.Invoke("GetAccKind", new object[] {
                        accYear,
                        acmWordNum});
            return ((string[])(results[0]));
        }
        
        /// <remarks/>
        public void GetAccKindAsync(string accYear, string acmWordNum) {
            this.GetAccKindAsync(accYear, acmWordNum, null);
        }
        
        /// <remarks/>
        public void GetAccKindAsync(string accYear, string acmWordNum, object userState) {
            if ((this.GetAccKindOperationCompleted == null)) {
                this.GetAccKindOperationCompleted = new System.Threading.SendOrPostCallback(this.OnGetAccKindOperationCompleted);
            }
            this.InvokeAsync("GetAccKind", new object[] {
                        accYear,
                        acmWordNum}, this.GetAccKindOperationCompleted, userState);
        }
        
        private void OnGetAccKindOperationCompleted(object arg) {
            if ((this.GetAccKindCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.GetAccKindCompleted(this, new GetAccKindCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/GetAccCount", RequestNamespace="http://tempuri.org/", ResponseNamespace="http://tempuri.org/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string[] GetAccCount(string accYear, string acmWordNum, string accKind) {
            object[] results = this.Invoke("GetAccCount", new object[] {
                        accYear,
                        acmWordNum,
                        accKind});
            return ((string[])(results[0]));
        }
        
        /// <remarks/>
        public void GetAccCountAsync(string accYear, string acmWordNum, string accKind) {
            this.GetAccCountAsync(accYear, acmWordNum, accKind, null);
        }
        
        /// <remarks/>
        public void GetAccCountAsync(string accYear, string acmWordNum, string accKind, object userState) {
            if ((this.GetAccCountOperationCompleted == null)) {
                this.GetAccCountOperationCompleted = new System.Threading.SendOrPostCallback(this.OnGetAccCountOperationCompleted);
            }
            this.InvokeAsync("GetAccCount", new object[] {
                        accYear,
                        acmWordNum,
                        accKind}, this.GetAccCountOperationCompleted, userState);
        }
        
        private void OnGetAccCountOperationCompleted(object arg) {
            if ((this.GetAccCountCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.GetAccCountCompleted(this, new GetAccCountCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/GetAccDetail", RequestNamespace="http://tempuri.org/", ResponseNamespace="http://tempuri.org/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string[] GetAccDetail(string accYear, string acmWordNum, string accKind, string accCount) {
            object[] results = this.Invoke("GetAccDetail", new object[] {
                        accYear,
                        acmWordNum,
                        accKind,
                        accCount});
            return ((string[])(results[0]));
        }
        
        /// <remarks/>
        public void GetAccDetailAsync(string accYear, string acmWordNum, string accKind, string accCount) {
            this.GetAccDetailAsync(accYear, acmWordNum, accKind, accCount, null);
        }
        
        /// <remarks/>
        public void GetAccDetailAsync(string accYear, string acmWordNum, string accKind, string accCount, object userState) {
            if ((this.GetAccDetailOperationCompleted == null)) {
                this.GetAccDetailOperationCompleted = new System.Threading.SendOrPostCallback(this.OnGetAccDetailOperationCompleted);
            }
            this.InvokeAsync("GetAccDetail", new object[] {
                        accYear,
                        acmWordNum,
                        accKind,
                        accCount}, this.GetAccDetailOperationCompleted, userState);
        }
        
        private void OnGetAccDetailOperationCompleted(object arg) {
            if ((this.GetAccDetailCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.GetAccDetailCompleted(this, new GetAccDetailCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/GetByPrimaryKey", RequestNamespace="http://tempuri.org/", ResponseNamespace="http://tempuri.org/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string GetByPrimaryKey(string accYear, string acmWordNum, string accKind, string accCount, string accDetail) {
            object[] results = this.Invoke("GetByPrimaryKey", new object[] {
                        accYear,
                        acmWordNum,
                        accKind,
                        accCount,
                        accDetail});
            return ((string)(results[0]));
        }
        
        /// <remarks/>
        public void GetByPrimaryKeyAsync(string accYear, string acmWordNum, string accKind, string accCount, string accDetail) {
            this.GetByPrimaryKeyAsync(accYear, acmWordNum, accKind, accCount, accDetail, null);
        }
        
        /// <remarks/>
        public void GetByPrimaryKeyAsync(string accYear, string acmWordNum, string accKind, string accCount, string accDetail, object userState) {
            if ((this.GetByPrimaryKeyOperationCompleted == null)) {
                this.GetByPrimaryKeyOperationCompleted = new System.Threading.SendOrPostCallback(this.OnGetByPrimaryKeyOperationCompleted);
            }
            this.InvokeAsync("GetByPrimaryKey", new object[] {
                        accYear,
                        acmWordNum,
                        accKind,
                        accCount,
                        accDetail}, this.GetByPrimaryKeyOperationCompleted, userState);
        }
        
        private void OnGetByPrimaryKeyOperationCompleted(object arg) {
            if ((this.GetByPrimaryKeyCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.GetByPrimaryKeyCompleted(this, new GetByPrimaryKeyCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/GetByKind", RequestNamespace="http://tempuri.org/", ResponseNamespace="http://tempuri.org/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string[] GetByKind(string accYear, string accKind, string batch) {
            object[] results = this.Invoke("GetByKind", new object[] {
                        accYear,
                        accKind,
                        batch});
            return ((string[])(results[0]));
        }
        
        /// <remarks/>
        public void GetByKindAsync(string accYear, string accKind, string batch) {
            this.GetByKindAsync(accYear, accKind, batch, null);
        }
        
        /// <remarks/>
        public void GetByKindAsync(string accYear, string accKind, string batch, object userState) {
            if ((this.GetByKindOperationCompleted == null)) {
                this.GetByKindOperationCompleted = new System.Threading.SendOrPostCallback(this.OnGetByKindOperationCompleted);
            }
            this.InvokeAsync("GetByKind", new object[] {
                        accYear,
                        accKind,
                        batch}, this.GetByKindOperationCompleted, userState);
        }
        
        private void OnGetByKindOperationCompleted(object arg) {
            if ((this.GetByKindCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.GetByKindCompleted(this, new GetByKindCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        public new void CancelAsync(object userState) {
            base.CancelAsync(userState);
        }
        
        private bool IsLocalFileSystemWebService(string url) {
            if (((url == null) 
                        || (url == string.Empty))) {
                return false;
            }
            System.Uri wsUri = new System.Uri(url);
            if (((wsUri.Port >= 1024) 
                        && (string.Compare(wsUri.Host, "localHost", System.StringComparison.OrdinalIgnoreCase) == 0))) {
                return true;
            }
            return false;
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.7.2046.0")]
    public delegate void GetVw_GBCVisaDetailJSONCompletedEventHandler(object sender, GetVw_GBCVisaDetailJSONCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.7.2046.0")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class GetVw_GBCVisaDetailJSONCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal GetVw_GBCVisaDetailJSONCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public string Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((string)(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.7.2046.0")]
    public delegate void FillVouNoCompletedEventHandler(object sender, FillVouNoCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.7.2046.0")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class FillVouNoCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal FillVouNoCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public string Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((string)(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.7.2046.0")]
    public delegate void GetYearCompletedEventHandler(object sender, GetYearCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.7.2046.0")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class GetYearCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal GetYearCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public string[] Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((string[])(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.7.2046.0")]
    public delegate void GetAcmWordNumCompletedEventHandler(object sender, GetAcmWordNumCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.7.2046.0")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class GetAcmWordNumCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal GetAcmWordNumCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public string[] Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((string[])(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.7.2046.0")]
    public delegate void GetAccKindCompletedEventHandler(object sender, GetAccKindCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.7.2046.0")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class GetAccKindCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal GetAccKindCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public string[] Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((string[])(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.7.2046.0")]
    public delegate void GetAccCountCompletedEventHandler(object sender, GetAccCountCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.7.2046.0")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class GetAccCountCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal GetAccCountCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public string[] Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((string[])(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.7.2046.0")]
    public delegate void GetAccDetailCompletedEventHandler(object sender, GetAccDetailCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.7.2046.0")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class GetAccDetailCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal GetAccDetailCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public string[] Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((string[])(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.7.2046.0")]
    public delegate void GetByPrimaryKeyCompletedEventHandler(object sender, GetByPrimaryKeyCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.7.2046.0")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class GetByPrimaryKeyCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal GetByPrimaryKeyCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public string Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((string)(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.7.2046.0")]
    public delegate void GetByKindCompletedEventHandler(object sender, GetByKindCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.7.2046.0")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class GetByKindCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal GetByKindCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public string[] Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((string[])(this.results[0]));
            }
        }
    }
}

#pragma warning restore 1591