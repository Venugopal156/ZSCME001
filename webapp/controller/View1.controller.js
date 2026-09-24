sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/export/Spreadsheet",
    "sap/m/MessageToast",
    "sap/m/MessageBox",
    "sap/ui/model/Filter",
    "sap/m/SearchField",
    "sap/ui/model/FilterOperator",
    "sap/ui/core/format/DateFormat",
    "../mylibrary"
], (Controller, Spreadsheet, MessageToast, MessageBox, Filter, SearchField, FilterOperator, DateFormat, MyLibrary) => {
    "use strict";

    return Controller.extend("com.zscme001.zscme001.controller.View1", {
        onInit() {


           
            var oModel = this.getView().getModel();
            this.getView().setModel(oModel, "oProgramModel");
            this.oResourceBundle = this.getOwnerComponent().getModel("i18n").getResourceBundle();
          

            var oInpValueModel = new sap.ui.model.json.JSONModel({
                SoldToPartyCSV: "",
                VendorCSV: "",
                DCATCSV: "",
                MaterialCSV: "",
                CustomerMasterMaterialCSV: ""
            });
            this.getView().setModel(oInpValueModel, "oInpValueModel");
            var oResponseModel = new sap.ui.model.json.JSONModel()
            this.getView().setModel(oResponseModel, "oResponseModel");

            //Model for Customer value help
            var oVHCustModel = new sap.ui.model.json.JSONModel();
            this.getView().setModel(oVHCustModel, "oVHCustModel");
            //Model for Vendor value help
            var oVHVendorModel = new sap.ui.model.json.JSONModel();
            this.getView().setModel(oVHVendorModel, "oVHVendorModel");
            //Model for Dcat value help
            var oDcatCustModel = new sap.ui.model.json.JSONModel();
            this.getView().setModel(oDcatCustModel, "oDcatCustModel");
            //Model for Storage Location value help
            var oStorageModel = new sap.ui.model.json.JSONModel();
            this.getView().setModel(oStorageModel, "oStorageModel");
            //Model for Material value help
            var oMaterialModel = new sap.ui.model.json.JSONModel();
            this.getView().setModel(oMaterialModel, "oMaterialModel");

            this.onVHLoaddata("I_Customer_VH", "oVHCustModel");
            this.onVHLoaddata("I_Supplier_VH", "oVHVendorModel");
            this.onVHLoaddata("ZMI_KNMT_DCAT_VH", "oDcatCustModel");
            this.onVHLoaddata("I_ProductVH", "oStorageModel");
            this.onVHLoaddata("I_CustomerMaterialValueHelp", "oMaterialModel");
            this._refreshData();
        },

        //Loading the valuehelp data
        onVHLoaddata: function (sEntitySet, sModel) {
            var that = this;
            var oModel = that.getOwnerComponent().getModel();
            oModel.read("/" + sEntitySet, {
                success: function (oData) {
                    var oJsonModel = that.getView().getModel(sModel)
                    oJsonModel.setData({
                        results: oData.results
                    });
                }
               
            })
        },

        //Upload Functionality

        onUpload: function () {

            const oFileUploader = this.byId("fileUploader");
            const oInput = oFileUploader.getDomRef("fu");
            const oFile = oInput && oInput.files && oInput.files[0];

            // =====================================================
            // FILE SELECTION CHECK
            // =====================================================
            if (!oFile) {
                MessageBox.warning("XLSXファイルを選択してください。");
                return;
            }

            // =====================================================
            // FILE TYPE CHECK
            // =====================================================
            if (!oFile.name.toLowerCase().endsWith(".xlsx")) {
                MessageBox.error("XLSXファイルのみ許可されています。");
                return;
            }

            // =====================================================
            // EMPTY FILE CHECK
            // =====================================================
            if (oFile.size === 0) {
                MessageBox.warning(
                    "アップロードするファイルが空です。データを入力してください。"
                );
                return;
            }

            const oModel = this.getView().getModel();
            const sFileName = oFile.name;
            const sFileType = oFile.type;

            // =====================================================
            // READ AND VALIDATE XLSX
            // =====================================================
            this._validateExcelFile(oFile)
                .then((bIsValid) => {

                    if (!bIsValid) {
                        return;
                    }

                    // =================================================
                    // FILE IS VALID - START UPLOAD
                    // =================================================
                    this.uploadkey = "X";
                    this.getView().setBusy(true);

                    return this._readFileAsBase64(oFile)
                        .then((sBase64Csv) => {

                            const oPayload = {
                                FileName: sFileName,
                                MimeType: sFileType,
                                FileContent: sBase64Csv
                            };

                            oModel.create("/UploadRequestSet", oPayload, {

                                success: (oData) => {

                                    this.getView().setBusy(false);
                                    oFileUploader.clear();

                                    const sUploadFile =
                                        oData && oData.UploadFile
                                            ? oData.UploadFile
                                            : "アップロードが完了しました";

                                    MessageBox.success(
                                        "ファイルをアップロード: " + sUploadFile
                                    );

                                    this.UploadFile = oData.FileContent;
                                    this.SubmittedBy = oData.UserName;
                                    this.Jobid = oData.JobId;

                                    delete oData.__metadata;

                                    this._refreshData(this.SubmittedBy);

                                    const oResponseModel =
                                        this.getView().getModel("oResponseModel");

                                    oResponseModel.setData({
                                        results: [oData]
                                    });
                                },

                                error: (oError) => {

                                    this.getView().setBusy(false);

                                    MessageBox.error(
                                        this._getErrorMessage(
                                            oError,
                                            "アップロードに失敗しました"
                                        )
                                    );
                                }
                            });
                        });
                })
                .catch(() => {

                    this.getView().setBusy(false);

                    MessageBox.error(
                        "Excelファイルの読み込みに失敗しました。"
                    );
                });
        },
        _validateExcelFile: function (oFile) {

            return new Promise((resolve, reject) => {

                const oReader = new FileReader();

                oReader.onload = (oEvent) => {

                    try {

                        // =====================================================
                        // READ EXCEL FILE
                        // =====================================================

                        const aData = new Uint8Array(
                            oEvent.target.result
                        );

                        const oWorkbook = XLSX.read(
                            aData,
                            {
                                type: "array"
                            }
                        );

                        // =====================================================
                        // CHECK 1 - SHEET EXISTS
                        // =====================================================

                        if (
                            !oWorkbook.SheetNames ||
                            oWorkbook.SheetNames.length === 0
                        ) {

                           
                             MessageBox.warning(
                                "データが存在しないためアップロードに失敗しました。"
                            );

                            resolve(false);
                            return;
                        }

                        // =====================================================
                        // GET FIRST SHEET
                        // =====================================================

                        const sSheetName =
                            oWorkbook.SheetNames[0];

                        const oWorksheet =
                            oWorkbook.Sheets[sSheetName];

                        if (!oWorksheet) {

                           
                             MessageBox.warning(
                                "データが存在しないためアップロードに失敗しました。"
                            );
                          

                            resolve(false);
                            return;
                        }

                        // =====================================================
                        // CONVERT EXCEL INTO ROWS
                        // =====================================================

                        const aRows =
                            XLSX.utils.sheet_to_json(
                                oWorksheet,
                                {
                                    header: 1,
                                    defval: ""
                                }
                            );

                        console.log("Excel rows:", aRows);

                        // =====================================================
                        // CHECK 2 - NO ROWS
                        // =====================================================

                        if (
                            !aRows ||
                            aRows.length === 0
                        ) {

                           
                             MessageBox.warning(
                                "データが存在しないためアップロードに失敗しました。"
                            );

                            resolve(false);
                            return;
                        }

                        // =====================================================
                        // CHECK 3 - HEADER
                        // =====================================================

                        const aHeaders =
                            aRows[0] || [];

                        console.log(
                            "Excel headers:",
                            aHeaders
                        );

                        const bHeaderEmpty =
                            aHeaders.length === 0 ||
                            aHeaders.every(
                                (sHeader) => {
                                    return String(sHeader)
                                        .trim() === "";
                                }
                            );

                        if (bHeaderEmpty) {

                           
                             MessageBox.warning(
                                "データが存在しないためアップロードに失敗しました。"
                            );

                            resolve(false);
                            return;
                        }

                        // =====================================================
                        // CHECK 4 - DATA
                        // =====================================================

                        const aDataRows =
                            aRows.slice(1);

                        console.log(
                            "Excel data rows:",
                            aDataRows
                        );

                        // No rows after header
                        if (
                            aDataRows.length === 0
                        ) {

                          
                             MessageBox.warning(
                                "データが存在しないためアップロードに失敗しました。"
                            );

                            resolve(false);
                            return;
                        }

                        // =====================================================
                        // CHECK WHETHER ALL DATA CELLS ARE EMPTY
                        // =====================================================

                        const bContentEmpty =
                            aDataRows.every(
                                (aRow) => {

                                    if (
                                        !aRow ||
                                        aRow.length === 0
                                    ) {
                                        return true;
                                    }

                                    return aRow.every(
                                        (sValue) => {
                                            return String(sValue)
                                                .trim() === "";
                                        }
                                    );
                                }
                            );

                        // =====================================================
                        // DATA IS EMPTY
                        // =====================================================

                        if (bContentEmpty) {

                            
                             MessageBox.warning(
                                "データが存在しないためアップロードに失敗しました。"
                            );

                            resolve(false);
                            return;
                        }

                        // =====================================================
                        // VALID FILE
                        // =====================================================

                        resolve(true);

                    } catch (oError) {

                        console.error(
                            "Excel validation error:",
                            oError
                        );

                        reject(oError);
                    }
                };

                // =========================================================
                // FILE READING ERROR
                // =========================================================

                oReader.onerror = () => {

                    reject(
                        new Error(
                            "Unable to read the Excel file."
                        )
                    );
                };

                // =========================================================
                // READ XLSX AS ARRAY BUFFER
                // =========================================================

                oReader.readAsArrayBuffer(oFile);
            });
        },

// Date Range Validation
        _validateDateRange: function () {
    const aDateFields = [
        this.byId("createdDateFrom"),
        this.byId("UpdatedDateFrom")
    ];
 
    for (const oDateRange of aDateFields) {
 
        const sValue = oDateRange.getValue();
 
        // Empty is allowed
        if (!sValue || sValue.trim() === "") {
            continue;
        }
 
        // DateRangeSelection automatically provides dateValue when
        // the entered value can be parsed as a valid date.
        const oDateValue = oDateRange.getDateValue();
        const oSecondDateValue = oDateRange.getSecondDateValue();
 
        // If user typed something invalid
        if (!oDateValue || isNaN(oDateValue.getTime())) {
 
            oDateRange.setValueState("Error");
            oDateRange.setValueStateText(
                "有効な日付を入力してください。"
            );
 
            MessageBox.error(
                "有効な日付を入力してください。"
            );
 
            return false;
        }
 
        // If second date exists, validate it too
        if (
            oSecondDateValue &&
            isNaN(oSecondDateValue.getTime())
        ) {
 
            oDateRange.setValueState("Error");
            oDateRange.setValueStateText(
                "有効な日付範囲を入力してください。"
            );
 
            MessageBox.error(
                "有効な日付範囲を入力してください。"
            );
 
            return false;
        }
 
        // Clear previous error
        oDateRange.setValueState("None");
        oDateRange.setValueStateText("");
    }
 
    return true;
},

//Second Tab Download Report Functionality

         onDownloadReport: function () {
            // =====================================================
    // DATE VALIDATION
    // =====================================================
    if (!this._validateDateRange()) {
        return;
    }
    var that = this;
    var oModel = that.getView().getModel("oInpValueModel");
     var oGetData = oModel.getData();
            var sPlant = oGetData.SoldToPartyCSV || "";
            var sProduct = oGetData.VendorCSV || "";
            var sDCat = oGetData.DCat || "";
            var sStorageLocationion = oGetData.MaterialCSV || "";
            var sMaster = oGetData.CustomerMasterMaterialCSV || "";
           

            // Setup Date Formatter (matches "yyyy/MM/dd" or adjust to "yyyyMMdd" based on backend requirement)
        var oDateFormat = DateFormat.getDateInstance({ pattern: "yyyyMMdd" });

        // Extract Date Objects directly from DateRangeSelection controls
        var oCreatedDateCtrl = this.byId("createdDateFrom"); // Replace with your actual View ID for Created Date
        var oUpdatedDateCtrl = this.byId("UpdatedDateFrom");     // Replace with your actual View ID for Updated Date

        var sCreatedDateFrom = "";
        var sCreatedDateTo = "";
        var sUpdatedDateFrom = "";
        var sUpdatedDateTo = "";

        if (oCreatedDateCtrl && oCreatedDateCtrl.getValue()) {
            if (oCreatedDateCtrl.getDateValue()) {
                sCreatedDateFrom = oDateFormat.format(oCreatedDateCtrl.getDateValue());
            }
            if (oCreatedDateCtrl.getSecondDateValue()) {
                sCreatedDateTo = oDateFormat.format(oCreatedDateCtrl.getSecondDateValue());
            }
        }

        if (oUpdatedDateCtrl && oUpdatedDateCtrl.getValue()) {
            if (oUpdatedDateCtrl.getDateValue()) {
                sUpdatedDateFrom = oDateFormat.format(oUpdatedDateCtrl.getDateValue());
            }
            if (oUpdatedDateCtrl.getSecondDateValue()) {
                sUpdatedDateTo = oDateFormat.format(oUpdatedDateCtrl.getSecondDateValue());
            }
        }


    var aFilters = [];

    // 1. Process Sold-to Party (SoldToPartyCSV) Multi-Values
    if (sPlant) {
        var aKunnrValues = sPlant.split(",").map(function (s) { return s.trim(); }).filter(Boolean);
        if (aKunnrValues.length > 0) {
            var aKunnrSubFilters = aKunnrValues.map(function (sValue) {
                return new sap.ui.model.Filter("SoldToPartyCSV", sap.ui.model.FilterOperator.EQ, sValue);
            });
            // Group multiple KUNNRs with 'OR' -> (KUNNR eq '1' or KUNNR eq '2')
            aFilters.push(new sap.ui.model.Filter({
                filters: aKunnrSubFilters,
                and: false // false creates an OR condition
            }));
        }
    }

    // 2. Process Vendor (VendorCSV) Multi-Values
    if (sProduct) {
        var aLifnrValues = sProduct.split(",").map(function (s) { return s.trim(); }).filter(Boolean);
        if (aLifnrValues.length > 0) {
            var aLifnrSubFilters = aLifnrValues.map(function (sValue) {
                return new sap.ui.model.Filter("VendorCSV", sap.ui.model.FilterOperator.EQ, sValue);
            });
            // Group multiple LIFNRs with 'OR' -> (LIFNR eq 'A' or LIFNR eq 'B')
            aFilters.push(new sap.ui.model.Filter({
                filters: aLifnrSubFilters,
                and: false // false creates an OR condition
            }));
        }
    }

    // 3. Process DCat (DCATCSV) Multi-Values
    if (sDCat) {
        var aDCatValues = sDCat.split(",").map(function (s) { return s.trim(); }).filter(Boolean);
        if (aDCatValues.length > 0) {
            var aDCatSubFilters = aDCatValues.map(function (sValue) {
                return new sap.ui.model.Filter("DCATCSV", sap.ui.model.FilterOperator.EQ, sValue);
            });
            // Group multiple ZDCATs with 'OR' -> (ZDCAT eq 'X' or ZDCAT eq 'Y')
            aFilters.push(new sap.ui.model.Filter({
                filters: aDCatSubFilters,
                and: false // false creates an OR condition
            }));
        }
    }
     // 4. Process Material  (MaterialCSV) Multi-Values

     if (sStorageLocationion) {
        var aMaterialValues = sStorageLocationion.split(",").map(function (s) { return s.trim(); }).filter(Boolean);
        if (aMaterialValues.length > 0) {
            var aMaterialSubFilters = aMaterialValues.map(function (sValue) {
                return new sap.ui.model.Filter("MaterialCSV", sap.ui.model.FilterOperator.EQ, sValue);
            });
            // Group multiple ZDCATs with 'OR' -> (ZDCAT eq 'X' or ZDCAT eq 'Y')
            aFilters.push(new sap.ui.model.Filter({
                filters: aMaterialSubFilters,
                and: false // false creates an OR condition
            }));
        }
    }
    // 4. Process Customer Material  (CustomerMasterMaterialCSV) Multi-Values

      if (sMaster) {
        var aCustomerMaterialValues = sMaster.split(",").map(function (s) { return s.trim(); }).filter(Boolean);
        if (aCustomerMaterialValues.length > 0) {
            var aCustomerMaterialSubFilters = aCustomerMaterialValues.map(function (sValue) {
                return new sap.ui.model.Filter("CustomerMasterMaterialCSV", sap.ui.model.FilterOperator.EQ, sValue);
            });
            // Group multiple ZDCATs with 'OR' -> (ZDCAT eq 'X' or ZDCAT eq 'Y')
            aFilters.push(new sap.ui.model.Filter({
                filters: aCustomerMaterialSubFilters,
                and: false // false creates an OR condition
            }));
        }
    }


      if (sCreatedDateFrom) {
            aFilters.push(new sap.ui.model.Filter("CreatedDateFrom", sap.ui.model.FilterOperator.EQ, sCreatedDateFrom));
        }
        if (sCreatedDateTo) {
            aFilters.push(new sap.ui.model.Filter("CreatedDateTo", sap.ui.model.FilterOperator.EQ, sCreatedDateTo));
        }
        if (sUpdatedDateFrom) {
            aFilters.push(new sap.ui.model.Filter("UpdatedDateFrom", sap.ui.model.FilterOperator.EQ, sUpdatedDateFrom));
        }
        if (sUpdatedDateTo) {
            aFilters.push(new sap.ui.model.Filter("UpdatedDateTo", sap.ui.model.FilterOperator.EQ, sUpdatedDateTo));
        }




    // 4. Force Parentheses Grouping using a Master Filter Array
    var aFinalFilterArray = [];
    if (aFilters.length > 0) {
        var oMasterFilter = new sap.ui.model.Filter({
            filters: aFilters,
            and: true // Combines distinct groups with 'AND' -> (KUNNR...) and (LIFNR...)
        });
        aFinalFilterArray.push(oMasterFilter);
    }

    // 5. Execute OData Request with the bundled filter array
    var oReadModel = that.getView().getModel();
    oReadModel.read("/DownloadRequestSet", {
        filters: aFinalFilterArray, // Using the grouped filter array
        success: function (oData) {
            if (oData && oData.results && oData.results.length > 0) {
                var oFILENAME = oData.results[0].FileName;
                var oFILECONTENT = oData.results[0].FileContent;
                
                if (oFILECONTENT === "E") {
                    MessageBox.warning(that.getView().getModel("i18n").getResourceBundle().getText("noFileFoundMessage"));
                    return;
                } else {
                    that._downloadResultFile(oFILECONTENT, oFILENAME);
                    return;
                }
            } else {
                
                
                MessageBox.warning("受注先コードか仕入先コードのいずれかを必ず入力してください");
                
            }
        }.bind(this),
        error: function (oError) {
            MessageBox.error(that._getErrorMessage(oError, "ダウンロードに失敗しました"));
        }.bind(this)
    });
},


        _readFileAsBase64: function (oFile) {
            return new Promise(function (resolve, reject) {
                var oReader = new FileReader();

                oReader.onload = function (oEvent) {
                    var sResult = oEvent.target.result || "";
                    var aParts = sResult.split(",");
                    resolve(aParts.length > 1 ? aParts[1] : "");
                };

                oReader.onerror = function () {
                    reject();
                };

                oReader.readAsDataURL(oFile);
            });
        },

        _getErrorMessage: function (oError) {
            var sDefaultMessage = "Upload failed.";

            try {
                var oResponse = JSON.parse(oError.responseText);
                if (oResponse.error && oResponse.error.message && oResponse.error.message.value) {
                    return oResponse.error.message.value;
                }
            } catch (e) {
                // ignore parse issue
            }

            return sDefaultMessage;
        },

        //third tab download report functionality

        onDownload: function () {
            var oTable = this.byId("idTable");
            var oResponseModel = this.getView().getModel("oResponseModel");

            var oDataModel = this.getOwnerComponent().getModel();
            var data = oResponseModel.getData();
            var that = this;

            var sUserName = "";

            this.getView().setBusy(true);

            oDataModel.read("/UploadRuntimeSet(UserName='" + sUserName + "')", {
                success: function (oData) {
                    var sFileName = oData.FileName;
                    var sMimeType = oData.MimeType;
                    var sFileContent = oData.FileContent;
                    that.getView().setBusy(false);

                    that._downloadResultFile(sFileContent, sFileName, sMimeType);
                   
                },
                error: function (oError) {
                    that.getView().setBusy(false);
                    MessageBox.error(that._getErrorMessage(oError));
                }
            });
        },

        _readFileAsBase64: function (oFile) {
            return new Promise(function (resolve, reject) {
                var oReader = new FileReader();

                oReader.onload = function (oEvent) {
                    var sResult = oEvent.target.result || "";
                    var aParts = sResult.split(",");
                    resolve(aParts.length > 1 ? aParts[1] : "");
                };

                oReader.onerror = function () {
                    reject();
                };

                oReader.readAsDataURL(oFile);
            });
        },

        //downloadresult file method
        _downloadResultFile: function (sResultFile, sUserName) {
 
            if (!sResultFile) {
                MessageBox.warning("No file content available.");
                return;
            }
 
            try {
                let sContent = sResultFile;
                const bIsBase64 = this._isBase64(sContent);
 
                /*
                 * UTF-8 BOM
                 * This was used in your original CSV download
                 * to handle Japanese characters correctly.
                 */
                const aBom = new Uint8Array([0xEF, 0xBB, 0xBF]);
 
                /*
                 * If backend FileContent is Base64,
                 * decode it first.
                 */
                if (bIsBase64) {
 
                    const sBinary = atob(sContent);
                    const aBytes = new Uint8Array(sBinary.length);
 
                    for (let i = 0; i < sBinary.length; i++) {
                        aBytes[i] = sBinary.charCodeAt(i);
                    }
 
                    /*
                     * Decode as UTF-8 so Japanese characters
                     * are not jumbled.
                     */
                    sContent = new TextDecoder("utf-8").decode(aBytes);
                }
 
                /*
                 * Remove BOM if it already exists in the content.
                 */
                if (sContent.charCodeAt(0) === 0xFEFF) {
                    sContent = sContent.substring(1);
                }
 
                /*
                 * Keep your original BOM + BlobParts logic.
                 *
                 * This preserves the UTF-8 CSV handling that
                 * was previously required for Japanese text.
                 */
                const aBlobParts = [];
 
                aBlobParts.push(aBom);
                aBlobParts.push(sContent);
 
                const oCsvBlob = new Blob(aBlobParts, {
                    type: "text/csv;charset=utf-8;"
                });
 
                /*
                 * Read the CSV content using SheetJS.
                 *
                 * The CSV is converted into an Excel workbook.
                 */
                const oWorkbook = XLSX.read(sContent, {
                    type: "string",
                    raw: true
                });
 
                /*
                 * Prepare Excel filename.
                 *
                 * Example:
                 * filename_0      -> filename_0.xlsx
                 * filename_0.csv  -> filename_0.xlsx
                 */
                let sExcelFileName = sUserName || "Report";
 
                // Remove existing extension, if present
                sExcelFileName = sExcelFileName.replace(/\.[^.]+$/, "");
 
                // Add .xlsx extension
                sExcelFileName = sExcelFileName + ".xlsx";
 
                console.log("Original filename:", sUserName);
                console.log("Final Excel filename:", sExcelFileName);
 
                /*
                 * Download REAL XLSX file.
                 */
                XLSX.writeFile(
                    oWorkbook,
                    sExcelFileName,
                    {
                        bookType: "xlsx",
                        compression: true
                    }
                );
 
                MessageBox.success(
                    "レポートのダウンロードが完了しました"
                );
 
            } catch (oError) {
 
                console.error("Excel download error:", oError);
 
                MessageBox.error(
                    "Excelファイルの作成に失敗しました。"
                );
            }
        },
        

        _isBase64: function (sValue) {
            if (!sValue || typeof sValue !== "string") {
                return false;
            }

            try {
                return btoa(atob(sValue)) === sValue.replace(/\s/g, "");
            } catch (e) {
                return false;
            }
        },

        _getErrorMessage: function (oError) {
            var sDefaultMessage = "Request failed.";

            try {
                var oResponse = JSON.parse(oError.responseText);
                if (oResponse.error && oResponse.error.message && oResponse.error.message.value) {
                    return oResponse.error.message.value;
                }
            } catch (e) {
                // ignore
            }

            return sDefaultMessage;
        },
       // Refresh Functionality
        onRefresh: function () {
            var that = this;
            var sUserName = this.UserName;
            var oModel = this.getView().getModel();
            oModel.read("/UploadRuntimeSet(UserName='" + sUserName + "')", {
                success: function (oData) {
                    that.getView().setBusy(false);

                    delete oData.__metadata;

                    var oResponseModel = that.getView().getModel("oResponseModel");

                    oResponseModel.setData({
                        results: [oData]
                    });

                    oResponseModel.refresh(true);

                    if (!oData || !oData.FileContent) {
                        MessageBox.warning("No result file found for UserName: " + sUserName);
                        return;
                    }

                },

                error: function (oError) {
                    that.getView().setBusy(false);
                    MessageBox.error(that._getErrorMessage(oError));
                }

            });
        },

        // Customer Value Help Method
        onValueHelpRequest: function () {
            
             this._sText = this.oResourceBundle.getText("SoldtoParty");
              this._fieldId = "I_Customer_VH";
             this._openPlantValueHelp("oVHCustModel", "Customer", "OrganizationBPName1");
        
           
        },
        // Supplier Value Help Method
        onValueHelpRequest2: function () {
             this._sText = this.oResourceBundle.getText("Vendor");
             this._fieldId = "I_Supplier_VH";
            this._openPlantValueHelp("oVHVendorModel", "Supplier", "SupplierName");
            
        },
        // DCAT Value Help Method
        onValueHelpRequest3: function () {
             this._sText = this.oResourceBundle.getText("DCAT");
             this._fieldId = "ZMI_THTSKNB_DCAT_VH";
            this._openPlantValueHelp("oDcatCustModel", "DCAT", "Description");
            
        },
        // Product Value Help Method
        onValueHelpRequest4: function () {
             this._sText = this.oResourceBundle.getText("Material");
                this._fieldId = "I_ProductVH";
            this._openPlantValueHelp("oStorageModel", "Product", "Product_Text");
         
        },
        // Customer Material Value Help Method
        onValueHelpRequest5: function () {
             this._sText = this.oResourceBundle.getText("CustomerMasterMaterial");
              this._fieldId = "I_CustomerMaterialValueHelp";
            this._openPlantValueHelp("oMaterialModel", "MaterialByCustomer", "MaterialDescriptionByCustomer");
           
        },

        //Value Help Logic
        _openPlantValueHelp: function (sModel, sProperty1, sProperty2, sEntitySet) {
            var that = this;
            that.Property = sProperty1;
            this.sProperty2 = sProperty2; // Store the entity set for later use
            var oResourceBundle = this.getView().getModel("i18n").getResourceBundle();
            var sLabel1 = oResourceBundle.hasText(sProperty1) ? oResourceBundle.getText(sProperty1) : sProperty1;
            var sLabel2 = oResourceBundle.hasText(sProperty2) ? oResourceBundle.getText(sProperty2) : sProperty2;

            if (!this._oPlantValueHelpDialog) {
                this._oPlantValueHelpDialog = sap.ui.xmlfragment("com.zscme001.zscme001.Fragment.plant", this);
                this.getView().addDependent(this._oPlantValueHelpDialog);
                

            }

             this._oPlantValueHelpDialog.setTitle(this._sText);
            //ensure the JSON Model exists on the view before loading/binding
            if (!this.getView().getModel(sModel)) {
                this.getView().setModel(new sap.ui.model.json.JSONModel({ results: [] }), sModel);
            }
            that._oPlantValueHelpDialog.setModel(
                this.getView().getModel(sModel), sModel
            );

            this.onVHLoaddata(sEntitySet, sModel);
            var oTable = new sap.m.Table({
                mode: "MultiSelect",
                selectionChange: this.onselectionChange.bind(this),
                columns: [
                    new sap.m.Column({ header: new sap.m.Label({ text: sLabel1 }) }),
                    new sap.m.Column({ header: new sap.m.Label({ text: sLabel2 }) })
                ]
            });

            oTable.bindItems({
                path: sModel + ">/results",
                template: new sap.m.ColumnListItem({
                    cells: [
                        new sap.m.Text({ text: "{" + sModel + ">" + sProperty1 + "}" }),
                        new sap.m.Text({ text: "{" + sModel + ">" + sProperty2 + "}" })
                    ]
                })
            });

            that._oSearchField = new SearchField({
                showSearchButton: true,
                search: function (oSearchEvent) {
                    this._triggerBackendSearch(oSearchEvent.getParameter("query"));
                }.bind(this)
            });

            var oFilterBar = this._oPlantValueHelpDialog.getFilterBar();
            if (oFilterBar) {
                oFilterBar.setBasicSearch(this._oSearchField);
            }

            that._oPlantValueHelpDialog.setTable(oTable);
            that._oPlantValueHelpDialog.setKey(sProperty1);
            that._oPlantValueHelpDialog.setDescriptionKey(sProperty2);

            that._oPlantValueHelpDialog.open();

            // that._oPlantValueHelpDialog.open();
        },


        
        //Search Functionality
        _triggerBackendSearch: function (sValue) {
            var that = this;
             if (sValue == "*") { sValue = undefined; }
            if (sValue && sValue.includes('*')) {
                sValue = sValue.replace(/\*/g, '');
            }
            var oTable = this._oPlantValueHelpDialog.getTable();
            var oBinding = oTable.getBinding("items");
            if (sValue) {
                var oFilter = new Filter(that.Property, FilterOperator.Contains, sValue);
                var oFilter1 = new Filter(that.sProperty2, FilterOperator.Contains, sValue);
                var oCombinedFilter = new Filter({
                    filters: [oFilter, oFilter1],
                    and: false // Set to true for AND logic; false for OR logic
                });
                oBinding.filter([oCombinedFilter]);
            } else {
                oBinding.filter([]);
            }
        },
       

        //Value help clolse
        onValueHelpClose: function (oEvent) {


            if (this._oPlantValueHelpDialog) {
                this._oPlantValueHelpDialog.close();
            }
        },

        //Value Helps search functionality
        onSearch: function (oEvent) {
            var that = this;
            var sValue = oEvent.getParameter("value");
            if (!sValue) {
                sValue = oEvent.getSource().getBasicSearchValue();
            }
            if (sValue == "*") { sValue = undefined; }
            if (sValue && sValue.includes('*')) {
                sValue = sValue.replace(/\*/g, '');
            }
            var oTable = this._oPlantValueHelpDialog.getTable();
            var oBinding = oTable.getBinding("items");
            if (sValue) {
                var oFilter = new Filter(that.Property, FilterOperator.Contains, sValue);
                var oFilter1 = new Filter(that.sProperty2, FilterOperator.Contains, sValue);
                var oCombinedFilter = new Filter({
                    filters: [oFilter, oFilter1],
                    and: false // Set to true for AND logic; false for OR logic
                });
                oBinding.filter([oCombinedFilter]);
            } else {
                oBinding.filter([]);
            }

        },
      
        onselectionChange: function (oEvent) {
             var oTable = oEvent.getSource();
            var aSelectedItems = oTable.getSelectedItems();
 
            var aSelectedValues = [];
 
            aSelectedItems.forEach(function (oItem) {
                var aCells = oItem.getCells();
 
                if (aCells && aCells.length > 0) {
                    aSelectedValues.push(aCells[0].getText());
                }
            });
 
            var sSelectedValue = aSelectedValues.join(",");

                var oInpValueModel = this.getView().getModel("oInpValueModel");
                if (this._fieldId === "I_Customer_VH") {
                    oInpValueModel.setProperty("/SoldToPartyCSV", sSelectedValue);
                } else if (this._fieldId === "I_Supplier_VH") {
                    oInpValueModel.setProperty("/VendorCSV", sSelectedValue);

                } else if (this._fieldId === "ZMI_THTSKNB_DCAT_VH") {
                    oInpValueModel.setProperty("/DCat", sSelectedValue);
                }
                else if (this._fieldId === "I_ProductVH") {
                    oInpValueModel.setProperty("/MaterialCSV", sSelectedValue);
                }
                else if (this._fieldId === "I_CustomerMaterialValueHelp") {
                    oInpValueModel.setProperty("/CustomerMasterMaterialCSV", sSelectedValue);
                }
           // }
        },
//Auto Refresh Method 
        _startAutoRefresh: function (userId) {

            if (this._refreshInterval) {
                clearInterval(this._refreshInterval);
            }

            this._refreshInterval = setInterval(() => {
                this._refreshData(userId);
            }, 100000);
        },
        _refreshData: function (userId) {
            var that = this;
            var oModel = that.getOwnerComponent().getModel();

            var oFilter = new sap.ui.model.Filter({
                path: "UserName",
                operator: sap.ui.model.FilterOperator.EQ,
                value1: userId
            });

            
            oModel.read("/UploadRuntimeSet", {
                success: (oData) => {

                    delete oData.__metadata;

                    var oResponseModel = that.getView().getModel("oResponseModel");

                
                    oResponseModel.setData(oData);
                    oResponseModel.refresh(true);
                    

                    var sStatus = oData.results[0].Status;
                    clearInterval(that._refreshInterval);
                    that._refreshInterval = null;
                    that.getView().byId("btnDownload").setEnabled(true);
                    that.getView().byId("btnUpload").setEnabled(true);
                
                },

                error: (oError) => {

                    clearInterval(that._refreshInterval);
                    that._refreshInterval = null;

                    that.getView().setBusy(false);

                    MessageBox.error(that._getErrorMessage(oError));
                }
            });
        },
       

        combineDateTime: function (sDate, sTime) {

            if (!sDate || !sTime) {
                return sDate || sTime || "";
            }
            return sDate + " " + sTime;
        }

    });
});