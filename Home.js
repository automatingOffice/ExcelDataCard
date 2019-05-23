(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#show-cards').click(createDataCards);
        });
    };


    function createDataCards() {
        Excel.run(function (ctx) {
            var excelDataSheet = ctx.workbook.worksheets.getItem('productData');
            var productRange = excelDataSheet.getRange('A1').getSurroundingRegion(); 
            productRange.load('values');
            ctx.sync().then(function () {
                var productDataArray = [];
                for (var i = 0; i < productRange.values.length; i++) {
                    var productItem = {
                        id: productRange.values[i][0],
                        product: productRange.values[i][1],
                        description: productRange.values[i][2],
                        imageUrl: productRange.values[i][3],
                        unitSales: productRange.values[i][4],
                        price: productRange.values[i][5]
                    };
                    productDataArray.push(productItem);
                }
                productDataArray = productDataArray.slice(1) 
                outputDataCard(productDataArray);
                return ctx.sync();
            });
        })
            .catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
    }

    function outputDataCard(productDataArray) {

        $('#product-listing').empty();

        productDataArray.forEach(function (item) {

            var numberFormat = '£' + (12345.67).toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,');

            //https://stackoverflow.com/questions/149055/how-can-i-format-numbers-as-dollars-currency-string-in-javascript

            var  myString = '<div class=\"card\" style=\"width: 18rem;">';
            myString = myString + '  <img src=\"https://picsum.photos/200/200" class=\"card-img-top\" alt=\"...\">';
            myString = myString + '<h5 class=\"card-title\">' + item.product + '</h5>';
            myString = myString + '<h6 class=\"card-subtitle mb-2 text-muted\">' + numberFormat + '</h6>';
            myString = myString + '    <p class=\"card-text\">' + item.description + '</p >';
            myString = myString + '</div>';
            myString = myString + '</div>';


            $('#product-listing').append(
                myString
            );
        });
    }





})();