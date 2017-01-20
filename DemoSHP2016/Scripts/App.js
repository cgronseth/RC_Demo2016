'use strict';

ExecuteOrDelayUntilScriptLoaded(initializePage, "sp.js");

function initializePage() {
    var context = SP.ClientContext.get_current();
    var user = context.get_web().get_currentUser();

    var urlLista = "https://redcomcibernetico.sharepoint.com/sites/desarrolloapps/AusenciasAPP";
    var nombreLista = "2017_Ausencias";

    var addinweburl;

    // Main
    $(document).ready(function () {
        getUserName();
        getVacaciones();
    });

    function getUserName() {
        context.load(user);
        context.executeQueryAsync(function () {
            $('#userName').text('Hola ' + user.get_title());
        },
        function (sender, args) {
            alert('Error:' + args.get_message());
        });
    }

    var hostweburl;
    var addinweburl;

    function getVacaciones() {
        hostweburl = decodeURIComponent(getQueryStringParameter("SPHostUrl"));
        addinweburl = decodeURIComponent(getQueryStringParameter("SPAppWebUrl"));

        document.getElementById("web3").innerHTML = "Url lista: " + urlLista;
        document.getElementById("web1").innerHTML = "Url Host: " + hostweburl;
        document.getElementById("web2").innerHTML = "Url Addin: " + addinweburl;

        var scriptbase = hostweburl + "/_layouts/15/";
        $.getScript(scriptbase + "SP.RequestExecutor.js", getVacacionesPaso2);
    }

    function getVacacionesPaso2() {
        var executor = new SP.RequestExecutor(addinweburl);
        executor.executeAsync(
            {
                url: addinweburl + "/_api/SP.AppContextSite(@target)/web/lists/getbytitle('" + nombreLista + "')/items?@target='" + urlLista + "'",
                method: "GET",
                headers: { "Accept": "application/json; odata=verbose" },
                success: function (data) {
                    document.getElementById("lista").innerHTML = "OK!";
                },
                error: function (data, errorCode, errorMessage) {
                    document.getElementById("error").innerHTML = "Error code: " + errorCode + "<br/>" + errorMessage + " :" + data.body;
                }
            }
        );
    }

    function getQueryStringParameter(paramToRetrieve) {
        var params = document.URL.split("?")[1].split("&");
        var strParams = "";
        for (var i = 0; i < params.length; i = i + 1) {
            var singleParam = params[i].split("=");
            if (singleParam[0] == paramToRetrieve)
                return singleParam[1];
        }
    }
}
