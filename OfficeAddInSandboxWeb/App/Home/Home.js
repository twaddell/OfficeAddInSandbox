/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#searchBtn').click(onSearchClick);
        });
    };

    var searchRanges = [];

    function onSearchClick() {
        cleanUp(searchRanges);
        runSearch("Video")
            .then(function(results) {
                searchRanges = results;
                displaySearchResults(searchRanges);
            });
    };

    function runSearch(textToFind) {
        var items = [];
        return Word.run(function (context) {
            var options = Word.SearchOptions.newObject(context);
            options.matchWildCards = false;

            var rangesFind = context.document.body.search(textToFind, options);
            context.load(rangesFind, 'text,');
            return context.sync()
                .then(function () {
                    $.each(rangesFind.items, function (i) {
                            items.push(rangesFind.items[i]);
                            context.trackedObjects.add(rangesFind.items[i]);
                    });
                    return context.sync();
                });
        })
        .then(function () {
            return items;
        });
    };

    function displaySearchResults(ranges) {
        var resultsList = $('#searchResults');
        resultsList.html('');
        $.each(ranges, function(i) {
            $('<li/>')
                .text(ranges[i].text)
                .click({range: ranges[i]}, selectRange)
                .appendTo(resultsList);
        });
    };

    function selectRange(event) {
        var range = event.data.range;
        range.select();
        range.context.sync()
        .catch(function(error) {
                console.log(error);
            });
    };

    function cleanUp(ranges) {
        $.each(ranges, function(i) {
            ranges[i].context.trackedObjects.remove(ranges[i]);
            ranges[i].context.sync();
        });
    };

})();