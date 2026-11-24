/**
 * Collections Messages RESTlet
 * Returns outbound email messages for a given month from saved search 7659.
 *
 * Deploy as a RESTlet in NetSuite:
 *   Customization → Scripting → Scripts → New
 *   Script Type: RESTlet
 *   Function: get
 *
 * GET parameters:
 *   month  — 1-12 (defaults to current month)
 *   year   — 4-digit year (defaults to current year)
 *
 * @NApiVersion 2.1
 * @NScriptType Restlet
 * @NModuleScope SameAccount
 */
define(['N/search', 'N/log'], function (search, log) {

    var SAVED_SEARCH_ID = '7659';

    function get(params) {
        var today = new Date();
        var month = parseInt(params.month, 10) || (today.getMonth() + 1);
        var year  = parseInt(params.year,  10) || today.getFullYear();

        // Build date range for the requested month
        var startDate = new Date(year, month - 1, 1);
        var endDate   = new Date(year, month,     1);   // first day of next month

        function nsDate(d) {
            // NetSuite date string: MM/DD/YYYY
            return (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear();
        }

        try {
            var savedSearch = search.load({ id: SAVED_SEARCH_ID });

            // Append date filter — keeps any existing filters on the saved search
            var filters = savedSearch.filters;
            filters.push(search.createFilter({
                name:     'messagedate',
                operator: search.Operator.ONORAFTER,
                values:   nsDate(startDate)
            }));
            filters.push(search.createFilter({
                name:     'messagedate',
                operator: search.Operator.BEFORE,
                values:   nsDate(endDate)
            }));
            savedSearch.filters = filters;

            var results = [];
            var pagedData = savedSearch.runPaged({ pageSize: 1000 });

            pagedData.pageRanges.forEach(function (pageRange) {
                var page = pagedData.fetch({ index: pageRange.index });
                page.data.forEach(function (result) {
                    results.push({
                        customer: result.getText({ name: 'entity' })
                                  || result.getValue({ name: 'entity' }) || '',
                        entity_id: result.getValue({ name: 'entity' }) || '',
                        author:    result.getValue({ name: 'author' })  || '',
                        date:      result.getValue({ name: 'messagedate' }) || '',
                        subject:   result.getValue({ name: 'subject' }) || ''
                    });
                });
            });

            log.debug({ title: 'collections_messages', details: 'Returned ' + results.length + ' rows for ' + month + '/' + year });

            return JSON.stringify({ count: results.length, month: month, year: year, results: results });

        } catch (e) {
            log.error({ title: 'collections_messages error', details: e.message });
            return JSON.stringify({ error: e.message, count: 0, results: [] });
        }
    }

    return { get: get };
});
