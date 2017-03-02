/// <reference path="../App.js" />
var storage = window.localStorage;

var username;
var password;

var user_range;
var people_range;
var transaction_range;
var selectedCellCalculate;

var selectCellRowIndex;
var selectCellColumnIndex;
(function () {
    "use strict";

    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            user_range = "";
            people_range = "";
            transaction_range = "";
            selectedCellCalculate = "";

            //controllo se mi sono gia' loggato e ho le credenziali salvate
            if (storage.logged) {
                if (storage.logged == "true") {
                    //riaggiorno i valori salvati in storage
                    var u = storage.username;
                    var p = storage.password;
                    var l = storage.logged;

                    storage.username = u;
                    storage.password = p;
                    storage.logged = l;

                    //faccio il login automatico
                    $("#username").val(storage.username);
                    $("#password").val(storage.password);

                    login();
                }
            } else {
                //permetto all'utente di fare il login e/o registrarsi
                $('#logout').prop("disabled", true);

                $("#a_site").show();

                $('#get-people-from-selection').prop("disabled", true);
                $('#get-transaction-from-selection').prop("disabled", true);
                $('#get-cell-calculate').prop("disabled", true);

                $('#calculate-tips').prop("disabled", true);
                $('#reset-tips').prop("disabled", true);
            }

            $('#login').click(login);
            $('#logout').click(logout);

            $('#get-people-from-selection').click(getPeopleFromSelection);
            $('#get-transaction-from-selection').click(getTransactionFromSelection);
            $('#get-cell-calculate').click(getCellCalculate);
            $('#calculate-tips').click(calculatetips);
            $('#reset-tips').click(reset);
        });
    };

    function login() {
        username = $("#username").val();
        password = $("#password").val();

        $.ajax({
            url: 'https://www.spreadsheetspace.net/orchestrator/login?action=loginFromAddin',
            type: 'POST',
            data: null,
            headers: { 'X-Username': username, 'X-Password': password },
            success: function (data, textStatus, jqXHR) {
                app.showNotification('Logged-in');

                storage.username = username;
                storage.password = password;
                storage.logged = true;

                $('#logout').prop("disabled", false);

                $("#a_site").hide();

                $('#get-people-from-selection').prop("disabled", false);
                $('#get-transaction-from-selection').prop("disabled", false);
                $('#get-cell-calculate').prop("disabled", false);

                $('#calculate-tips').prop("disabled", false);
                $('#reset-tips').prop("disabled", false);
            },
            error: function (jqXHR, textStatus, errorThrown) {
                app.showNotification('error');

            }
        });
    }

    function logout() {
        //svuoto lo storage salvato
        storage.clear();

        //faccio il reset delle caselle di testo
        $("#username").val("");
        $("#password").val("");
        $("#input-people").val("");
        $("#input-transaction").val("");
        $("#input-calculate").val("");

        //mostro il link per registrarsi
        $("#a_site").show();

        //disabilito tutti i bottoni non utilizzabili
        $('#logout').prop("disabled", true);

        $('#get-people-from-selection').prop("disabled", true);
        $('#get-transaction-from-selection').prop("disabled", true);
        $('#get-cell-calculate').prop("disabled", true);

        $('#calculate-tips').prop("disabled", true);
        $('#reset-tips').prop("disabled", true);
    }

    function reset() {
        //faccio il reset di tutti i parametri salvati
        user_range = "";
        people_range = "";
        transaction_range = "";
        selectedCellCalculate = "";

        //faccio il reset delle caselle di testo
        $("#input-people").val("");
        $("#input-transaction").val("");
        $("#input-calculate").val("");
    }

    function getPeopleFromSelection() {
        Office.context.document.bindings.addFromPromptAsync(Office.BindingType.Matrix, { id: 'peopleMatrix' }, function (result) {
            //apro il prompt per fare la selezione delle celle per il range
            if (result.status == 'succeeded') {
                Excel.run(function (ctx) {
                    //carico l'address che potro' poi usare solo nella return ctx.sync
                    var binding = ctx.workbook.bindings.getItem("peopleMatrix");
                    var range = binding.getRange();
                    range.load("address");

                    return ctx.sync().then(function () {
                        Office.select("bindings#peopleMatrix", function onError() { }).getDataAsync(function (asyncResult) {
                            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                                console.log('Action failed. Error: ' + asyncResult.error.message);
                            } else {
                                //memorizzo i valori selezionati
                                people_range = asyncResult.value;
                            }
                        });

                        //utilizzo l'address caricato prima e ne faccio il display nella casella di input
                        $("#input-people").val(range.address);
                    }).catch(function (error) {
                        console.log('Error:', error.message);
                    });
                }).catch(function (error) {
                    console.log("Error: " + error);
                });
            } else {
                console.log('Error:', result.error.message);
            }
        });
    }

    function getTransactionFromSelection() {
        Office.context.document.bindings.addFromPromptAsync(Office.BindingType.Matrix, { id: 'transactionMatrix' }, function (result) {
            //apro il prompt per fare la selezione delle celle per il range
            if (result.status == 'succeeded') {
                Excel.run(function (ctx) {
                    //carico l'address che potro' poi usare solo nella return ctx.sync
                    var binding = ctx.workbook.bindings.getItem("transactionMatrix");
                    var range = binding.getRange();
                    range.load("address");

                    return ctx.sync().then(function () {
                        Office.select("bindings#transactionMatrix", function onError() { }).getDataAsync(function (asyncResult) {
                            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                                console.log('Action failed. Error: ' + asyncResult.error.message);
                            } else {
                                //memorizzo i valori selezionati
                                transaction_range = asyncResult.value;
                            }
                        });

                        //utilizzo l'address caricato prima e ne faccio il display nella casella di input
                        $("#input-transaction").val(range.address);
                    }).catch(function (error) {
                        console.log('Error:', error.message);
                    });
                }).catch(function (error) {
                    console.log("Error: " + error);
                });
            } else {
                console.log('Error:', result.error.message);
            }
        });
    }

    function getCellCalculate() {
        Office.context.document.bindings.addFromPromptAsync(Office.BindingType.Matrix, { id: 'cellMatrix' }, function (result) {
            //apro il prompt per fare la selezione delle celle per il range
            if (result.status == 'succeeded') {
                Excel.run(function (ctx) {
                    //carico l'address che potro' poi usare solo nella return ctx.sync
                    //inoltre carico rowIndex e columnIndex che mi serviranno per fare lo shit del range che otterro' dal ricalcolo
                    //infine carico rowCount e rowIndex che utilizzero' per controllare se ho selezionato una sola cella
                    var binding = ctx.workbook.bindings.getItem("cellMatrix");
                    var range = binding.getRange();
                    range.load("address");
                    range.load("rowIndex");
                    range.load("columnIndex");
                    range.load("rowCount");
                    range.load("columnCount");

                    return ctx.sync().then(function () {
                        if (range.rowCount == 1 && range.columnCount == 1) {
                            //memorizzo i valori caricati
                            selectedCellCalculate = range.address;
                            selectCellRowIndex = range.rowIndex;
                            selectCellColumnIndex = range.columnIndex

                            //utilizzo l'address caricato prima e ne faccio il display nella casella di input
                            $("#input-calculate").val(range.address);
                        } else {
                            app.showNotification('Error. You can select only one cell');
                        }

                    }).catch(function (error) {
                        console.log('Error:', error.message);
                    });
                }).catch(function (error) {
                    console.log("Error: " + error);
                });
            } else {
                console.log('Error:', result.error.message);
            }
        });
    }

    function calculatetips() {
        var user = [];
        var tmpUser = [];
        tmpUser.push(username);
        user.push(tmpUser);
 
        var dataToSend = {};
        var dataToSendJSON;
        var url = "https://www.spreadsheetspace.net/SSSServices/rest/SSSServices/recalculate"


        if (people_range == "" || transaction_range == "") {
            app.showNotification('Error. Select Transaction and/or People range before');
        } else {
            if (selectedCellCalculate == "") {
                app.showNotification('Error. You must select the cell that you want to show the Recalculation');
            } else {
                //creo il JSON da inviare in accordo con le specifiche del server 
                dataToSend = {
                    "user": user,
                    "people": people_range,
                    "transactions": transaction_range
                }

                dataToSendJSON = JSON.stringify(dataToSend);

                $.ajax({
                    url: url,
                    type: 'POST',
                    data: dataToSendJSON,
                    contentType: 'text/xml',
                    success: function (data, status, jqXHR) {
                        Excel.run(function (ctx) {
                            //a partire dall'address salvato in precedenza, ricavo la cella di partenza in cui copiare il risultato
                            var index = selectedCellCalculate.indexOf("!") + 1;
                            var wb = selectedCellCalculate.substring(0, index-1);
                            var cell = selectedCellCalculate.substring(index);

                            //utilizzando rowIndex, columnIndex e le dimensioni del dato ottenuto ricavo l'address dell'ultima cella in cui andro' ad incollare i dati
                            var sheet = ctx.workbook.worksheets.getItem(wb);
                            var firstCellRange = sheet.getRange(cell + ":" + cell);
                            var firstCell = sheet.getCell(selectCellRowIndex, selectCellColumnIndex);
                            var lastCell = sheet.getCell(selectCellRowIndex + data.length - 1, selectCellColumnIndex + data[0].length - 1);
                            lastCell.load('addressLocal');

                            return ctx.sync().then(function () {
                                //creo il range con i dati calcolati prima ed incollo il risultato ottenuto dal ricalcolo
                                var range = sheet.getRange(cell + ":" + lastCell.addressLocal)
                                range.values = data;

                                app.showNotification('SpreadSheetSpace Services: recalculation completed.');
                            }).catch(function (error) {
                                console.log('Error:', error);
                            });
                        }).catch(function (error) {
                            console.log("Error: " + error);
                        });
                    },
                    error: function (jqXHR, status, errorThrown) {
                        app.showNotification('Error. Something went wrong. Maybe you should check your input data...');
                    }

                });
            }
        }
    }
})();