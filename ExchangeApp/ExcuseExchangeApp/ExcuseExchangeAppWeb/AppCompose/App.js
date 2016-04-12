/* Common app functionality */

var app = (function () {
    "use strict";

    var app = {};

    // Common initialization function (to be called from each page)
    app.initialize = function () {
    	$('#generate').click(function () {
    		var name = $('#name').val();
    		$('#put').removeAttr("disabled");
    		$('#excuse').text(Generator.generate(name));
    	});

    	$('#put').click(function () {
    		var textToInsert = $('#excuse').text();
    		Office.context.mailbox.item.body.setSelectedDataAsync(
			  textToInsert,
			  { coercionType: Office.CoercionType.Text },
			  function (asyncResult) {
			  	if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
			  		app.showNotification("Success", "\"" + textToInsert + "\" inserted successfully.");
			  	}
			  	else {
			  		app.showNotification("Error", "Failed to insert \"" + textToInsert + "\": " + asyncResult.error.message);
			  	}
			  });
    	});
    };

    return app;
})();