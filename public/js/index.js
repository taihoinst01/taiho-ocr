var excelParams = []; // 엑셀 작업에 필요한 파라미터
var checkNum = 0;

$(function () {
    var dt = new Date();
    var current = dt.getFullYear() + '-' + ((dt.getMonth() + 1 < 10) ? '0' + (dt.getMonth() + 1) : dt.getMonth() + 1) + '-' + ((dt.getDate() + 1 < 10) ? '0' + (dt.getDate()) : dt.getDate())
    $('#fdate').val(current);
    $('#tdate').val(current);
    excelSaveEvent();
    dateEvent();
    getFiles();
    uploadImg();
    ocrApi();
});

function getFiles() {
    $.ajax({
        url: '/getFiles',
        type: 'get',
        contentType: 'application/json; charset=UTF-8',
        success: function (data) {
            for (var i = 0; i < data.files.length; i++) {
                var itemHTML = '<tr>';
                itemHTML += '<td><input name="imgCheck" type="checkbox" value="' + data.files[i].fileName.split('/')[2] + '"> </td>';
                itemHTML += '<td>' + data.files[i].fileName.split('/')[2] + '</td>';
                itemHTML += '<td>' + data.files[i].time.split('T')[0] + '</td>';
                itemHTML += '</tr>';
                $('#result').append(itemHTML);
            }
            if (data.count != 0) {
                $('.mtm10 > strong').text(data.count + 1);
            } else {
                $('.mtm10 > strong').text(data.count);
            }
            checkBoxEvent();

        },
        error: function (err) {
            console.log(err);
        }
    });
}

function uploadImg() {
    $('#uploadFile').change(function () {
        var extArr = ['jpg', 'jpeg', 'png', 'gif', 'bmp'];
        var uploadFile = $(this).val();
        var filename = uploadFile.split('/').pop().split('\\').pop();

        var lastDot = uploadFile.lastIndexOf('.');
        var fileExt = uploadFile.substring(lastDot + 1, uploadFile.length).toLowerCase();
        if ($.inArray(fileExt, extArr) != -1 && $(this).val() != '') {
            var formData = new FormData($('#uploadForm')[0]);
            $.ajax({
                url: '/upload',
                type: 'POST',
                data: formData,
                processData: false,
                contentType: false,
                success: function (msg) {
                    location.reload();
                }
            });
        } else {
            $(this).val('');
            alert('파일 형식이 올바르지 않습니다.');
        }
    });

    $('#uploadBtn').click(function () {
        $('#uploadFile').click();
    });
}

function checkBoxEvent() {
    $('input[type="checkbox"]').change(function (e) {
        var checkCount = 0;
        if ($(e.target).val() == 'all') {
            if ($(e.target).is(':checked')) {
                $("input[type=checkbox]").prop("checked", true);
            } else {
                $("input[type=checkbox]").prop("checked", false);
            }
        }

        $("input[type=checkbox]:checked").each(function (i, el) {
            if (el.value != 'all') {
                checkCount++;
            }
        });
        $('.mtp10 > strong').text(checkCount);
        checkNum = checkCount;
    });
}

function dateEvent() {
    $('.date').click(function (e) {
        $('.date').each(function (i, el) {
            $(el).removeClass('on');
        });
        $(e.target).addClass('on');
        type = $(e.target).attr('alt');
        dateHandler(type);
    });
    $('#fdate,#tdate').change(function () {
        dateHandler('condition');
    });
}

function dateHandler(type) {
    var dateInterval;

    dateInterval = conditionalSearch(type);
    $.ajax({
        url: '/getFiles',
        type: 'get',
        contentType: 'application/json; charset=UTF-8',
        success: function (data) {
            $('#result').html('');
            var itemCount = 0;
            for (var i = 0; i < data.files.length; i++) {
                var itemTime = data.files[i].time.split('T')[0].split('-');
                var itemDate = itemTime[0] + itemTime[1] + itemTime[2];
                if (dateInterval == 0 || (dateInterval[0] >= itemDate && itemDate >= dateInterval[1])) {
                    var itemHTML = '<tr>';
                    itemHTML += '<td><input name="imgCheck" type="checkbox" value="' + data.files[i].fileName.split('/')[2] + '"> </td>';
                    itemHTML += '<td>' + data.files[i].fileName.split('/')[2] + '</td>';
                    itemHTML += '<td>' + data.files[i].time.split('T')[0] + '</td>';
                    itemHTML += '</tr>';
                    itemCount++;
                    $('#result').append(itemHTML);
                }
            }
            $('.mtm10 > strong').text(itemCount);
            checkBoxEvent();

        },
        error: function (err) {
            console.log(err);
        }
    });
}

function conditionalSearch(type) {
    var startDate = dateConvert(new Date());
    var endDate;
    var result = [];

    if (type == 'all') {
    } else if (type == 'today') {
        endDate = dateConvert(new Date());
        result.push(startDate);
        result.push(endDate);
    } else if (type == 'week') {
        var d = new Date();
        var weekDate = new Date().getTime() - (7 * 24 * 60 * 60 * 1000);
        d.setTime(weekDate);
        endDate = dateConvert(d);
        result.push(startDate);
        result.push(endDate);
    } else if (type == 'month') {
        var d = new Date();
        var monthDate = new Date().getTime() - (30 * 24 * 60 * 60 * 1000);
        d.setTime(monthDate);
        endDate = dateConvert(d);
        result.push(startDate);
        result.push(endDate);
    } else {
        var tdate = $('#tdate').val();
        var fdate = $('#fdate').val();
        result.push(tdate.split('-')[0] + tdate.split('-')[1] + tdate.split('-')[2]);
        result.push(fdate.split('-')[0] + fdate.split('-')[1] + fdate.split('-')[2]);
    }
    return result;
}

function dateConvert(d) {
    return '' + d.getFullYear() + (((d.getMonth() + 1) < 10) ? '0' + (d.getMonth() + 1) : (d.getMonth() + 1)) + ((d.getDate() < 10) ? '0' + d.getDate() : d.getDate());
}

function ocrApi() {
    $('#ocrBtn').click(function () {
        var checkCount = 0;
        $('.tex01').text('분석 상세 내용');
        excelParams = [];
        $('.tnum01').text('0');
        $('.t01').text('0');
        $('.t04').text('0');
        $("input[type=checkbox]:checked").each(function (i, el) {
            if (el.value != 'all') {
                processImage($(el).val());
                checkCount++;
            }
        });

        if (checkCount > 0) {
            $("input[type=checkbox]").prop("checked", false);
            $('.mtp10 > strong').text('0');
            //analysisImg();
        } else {
            alert('분석할 문서를 선택해주세요');
        }
    });
}

//OCR API 호출
function processImage(fileName) {
    $('#dataForm').html('');
    var subscriptionKey = "fedbc6bb74714bd78270dc8f70593122";
    var uriBase = "https://westus.api.cognitive.microsoft.com/vision/v1.0/ocr";

    // Request parameters.
    var params = {
        "language": "ko",
        "detectOrientation": "true",
    };

    // image url
    var sourceImageUrl = 'http://kr-ocr.azurewebsites.net/uploads/' + fileName;
    console.log(sourceImageUrl);

    // Perform the REST API call.
    $.ajax({
        url: uriBase + "?" + $.param(params),

        // Request headers.
        beforeSend: function (jqXHR) {
            jqXHR.setRequestHeader("Content-Type", "application/json");
            jqXHR.setRequestHeader("Ocp-Apim-Subscription-Key", subscriptionKey);
        },

        type: "POST",

        // Request body.
        data: '{"url": ' + '"' + sourceImageUrl + '"}',
    }).done(function (data) {
        //console.log(data);
        $('.tnum01').text((Number($('.tnum01').text()) + 1));
        ocrDataProcessing(data.regions, fileName);

    }).fail(function (jqXHR, textStatus, errorThrown) {
        var errorString = (errorThrown === "") ? "Error. " : errorThrown + " (" + jqXHR.status + "): ";
        errorString += (jqXHR.responseText === "") ? "" : (jQuery.parseJSON(jqXHR.responseText).message) ?
            jQuery.parseJSON(jqXHR.responseText).message : jQuery.parseJSON(jqXHR.responseText).error.message;
        alert(errorString);
    });
};

function analysisImg(type) {
    //console.log(excelParams)
    var params = { 'data': excelParams };

    $.ajax({
        url: '/uploadExcel',
        type: 'post',
        datatype: "json",
        data: JSON.stringify(params),
        contentType: 'application/json; charset=UTF-8',
        success: function (data) {
            //$('#successExcelName').val(data.successExcelName);
            //$('#failExcelName').val(data.failExcelName);
            setTimeout(function () {
                if (type == 'success') {
                    window.location = '/downloadExcel?fileName=' + data.successExcelName;
                } else {
                    window.location = '/downloadExcel?fileName=' + data.failExcelName;
                }
            }, 1000);           
        },
        error: function (err) {
            console.log(err);
        }
    });
}

function ocrDataProcessing(regions, fileName) {
    //console.log(regions);
    var lineText = []
    var docDate, companyName, contractName, email,
        price1, price2, price3,
        totPrice, detail, etc,
        bookingNum, shipper, shipperRegion, cosignee, cosigneeRegion,
        documentType;
    for (var i = 0; i < regions.length; i++) {
        for (var j = 0; j < regions[i].lines.length; j++) {
            var item = '';
            for (var k = 0; k < regions[i].lines[j].words.length; k++) {
                item += regions[i].lines[j].words[k].text + ' ';
            }
            lineText.push({ 'location': regions[i].lines[j].boundingBox, 'text': item });
        }
    }
    console.log(lineText);

    for (var i = 0; i < lineText.length; i++) {

        if (lineText[i].location == '1448,419,391,31' && lineText[i].text.trim() == 'ACCOUNT INVOICE') { // ACCOUNT INVOIC 양식
            documentType = 1;
            for (var j = 0; j < lineText.length; j++) {
                switch (lineText[j].location) {
                    case '2056,728,190,30':
                    case '2056,728,201,30':
                        docDate = lineText[j].text.trim();
                        break;
                    case '426,598,467,33':
                        companyName = lineText[j].text.trim();
                        break;
                    case '671,1046,749,33':
                    case '671,1046,55,26':
                    case '671,1046,111,33':
                    case '671,1046,409,33':
                    case '670,1046,290,33':
                    case '670,1046,296,33':
                    case '671,1046,205,33':
                    case '680,1046,635,33':
                        contractName = lineText[j].text.trim();
                        break;
                    case '1944,597,483,29':
                        email = lineText[j].text.trim();
                        break;
                    case '2792,1472,153,31':
                    case '2841,1472,104,26':
                    case '2830,1515,115,26':
                    case '2841,1472,99,26':
                    case '2844,1472,96,26':
                    case '2813,1472,132,31':
                    case '2812,1472,133,31':
                    case '2793,1472,151,31':
                        price1 = lineText[j].text.trim();
                        break;
                    case '2849,1515,96,26':
                    case '2880,1515,65,26':
                    case '2830,1558,115,26':
                    case '2848,1557,107,33':
                    case '2868,1515,77,26':
                    case '2849,1515,95,26':
                        price2 = lineText[j].text.replace('(', '').replace(')', '').trim();
                        break;
                    case '2801,1558,144,30':
                    case '2782,1558,163,30':
                    case '2801,1554,144,30':
                    case '2849,1558,91,26':
                    case '2849,1558,96,26':
                    case '2849,1515,96,26':
                        price3 = lineText[j].text.trim();
                        break;
                    case '2803,1642,151,30':
                    case '2782,1642,163,30':
                    case '2793,1646,152,30':
                    case '2830,1643,115,25':
                    case '2844,1649,96,25':
                    case '2850,1642,104,26':
                    case '2870,1642,79,26':
                    case '2848,1642,107,32':
                    case '2822,1642,132,30':
                    case '2822,1642,127,30':
                    case '2802,1642,147,30':
                    case '2867,1642,88,32':
                        totPrice = lineText[j].text.replace('(', '').replace(')', '').trim();
                        break;
                }
            }
        } else if (lineText[i].location == '3137,1128,564,57' && lineText[i].text.trim() == 'ACCOUNT INVOICE') { // 축소한 ACCOUNT INVOIC 양식
            documentType = 1;
            for (var j = 0; j < lineText.length; j++) {
                switch (lineText[j].location) {
                    case '4003,1583,274,44':
                        docDate = lineText[j].text.trim();
                        break;
                    case '1657,1339,676,63':
                        companyName = lineText[j].text.trim();
                        break;
                    case '1994,1975,1081,72':
                        contractName = lineText[j].text.trim();
                        break;
                    case '3848,1396,696,56':
                        email = lineText[j].text.trim();
                        break;
                    case '5107,2654,151,39':
                        price1 = lineText[j].text.trim();
                        break;
                    case '5163,2714,93,39':
                        price2 = lineText[j].text.trim();
                        break;
                    case '5117,2774,137,38':
                        price3 = lineText[j].text.trim();
                        break;
                    case '5114,2892,151,39':
                        totPrice = lineText[j].text.trim();
                        break;
                }
            }
        } else if (lineText[i].location == '1253,492,251,22' && lineText[i].text.trim() == 'ACCOUNT INVOICE') { // 기울어진 ACCOUNT INVOIC 양식
            documentType = 1;
            for (var j = 0; j < lineText.length; j++) {
                switch (lineText[j].location) {
                    case '1643,690,122,20':
                        docDate = lineText[j].text.trim();
                        break;
                    case '598,608,301,22':
                        companyName = lineText[j].text.trim();
                        break;
                    case '756,894,480,23':
                        contractName = lineText[j].text.trim();
                        break;
                    case '1571,606,311,22':
                        email = lineText[j].text.trim();
                        break;
                    case '2146,1166,67,18':
                        price1 = lineText[j].text.trim();
                        break;
                    case '2170,1193,43,18':
                        price2 = lineText[j].text.trim();
                        break;
                    case '2151,1220,62,18':
                        price3 = lineText[j].text.trim();
                        break;
                    case '2152,1275,68,17':
                        totPrice = lineText[j].text.trim();
                        break;
                }
            }
        } else if (lineText[i].location == '2586,713,998,57' && lineText[i].text.trim() == 'STATEMENT OF ACCOUNT') { // STATEMENT OF ACCOUNT 양식
            documentType = 1;
            for (var j = 0; j < lineText.length; j++) {
                switch (lineText[j].location) {
                    case '592,1392,672,57':
                        docDate = lineText[j].text.replace('Created', '').trim();
                        break;
                    case '593,1584,1048,70':
                        companyName = lineText[j].text.trim();
                        break;
                    case '2609,2733,1241,58':
                        contractName = lineText[j].text.trim();
                        break;
                    case '592,5325,1179,72':
                        email = lineText[j].text.replace('Contact details:', '').trim();
                        break;
                    case '5237,3599,335,67':
                        price1 = lineText[j].text.trim();
                        break;
                    case '5343,3791,229,57':
                        price2 = lineText[j].text.trim();
                        break;
                    case '5280,3887,292,67':
                        price3 = lineText[j].text.trim();
                        break;
                    case '5237,4079,335,67':
                        totPrice = lineText[j].text.trim();
                        break;
                }
            }
        } else if (lineText[i].location == '2840,1093,851,49' && lineText[i].text.trim() == 'STATEMENT OF ACCOUNT') { // 축소한 STATEMENT OF ACCOUNT 양식
            documentType = 1;
            for (var j = 0; j < lineText.length; j++) {
                switch (lineText[j].location) {
                    case '1135,1662,575,48':
                        docDate = lineText[j].text.replace('Created', '').trim();
                        break;
                    case '1135,1821,896,58':
                        companyName = lineText[j].text.trim();
                        break;
                    case '2858,2782,1060,48':
                        contractName = lineText[j].text.trim();
                        break;
                    case '1135,4948,1008,60':
                        email = lineText[j].text.replace('Contact details:', '').trim();
                        break;
                    case '5103,3505,286,56':
                        price1 = lineText[j].text.trim();
                        break;
                    case '5194,3665,195,49':
                        price2 = lineText[j].text.trim();
                        break;
                    case '5140,3746,249,56':
                        price3 = lineText[j].text.trim();
                        break;
                    case '5103,3906,286,57':
                        totPrice = lineText[j].text.trim();
                        break;
                }
            }
        } else if (lineText[i].location == '2921,911,998,58' && lineText[i].text.trim() == 'STATEMENT OF ACCOUNT') { // 기울어진 STATEMENT OF ACCOUNT 양식
            documentType = 1;
            for (var j = 0; j < lineText.length; j++) {
                switch (lineText[j].location) {
                    case '926,1590,673,58':
                        docDate = lineText[j].text.replace('Created', '').trim();
                        break;
                    case '927,1782,1049,71':
                        companyName = lineText[j].text.trim();
                        break;
                    case '2943,2931,1242,59':
                        contractName = lineText[j].text.trim();
                        break;
                    case '926,5523,1180,73':
                        email = lineText[j].text.replace('Contact details:', '').trim();
                        break;
                    case '5572,3797,335,68':
                        price1 = lineText[j].text.trim();
                        break;
                    case '5677,3989,231,58':
                        price2 = lineText[j].text.trim();
                        break;
                    case '5615,4085,292,68':
                        price3 = lineText[j].text.trim();
                        break;
                    case '5571,4277,336,68':
                        totPrice = lineText[j].text.trim();
                        break;
                }
            }
        } else if (lineText[i].location == '461,130,333,38' && lineText[i].text.trim() == 'PACKING LIST') {
            documentType = 1;
            for (var j = 0; j < lineText.length; j++) {
                switch (lineText[j].location) {
                    case '448,782,130,18':
                        docDate = lineText[j].text.replace('Created', '').trim();
                        break;
                    case '71,231,267,20':
                        companyName = lineText[j].text.split('CO.,')[0].trim();
                        break;
                }
            }
        }
        //KSO
        else if (lineText[i].location == '1221,336,799,87' && lineText[i].text.trim() == 'SHIPPING REQUEST') {   // 03
            documentType = 2;
            for (var j = 0; j < lineText.length; j++) {
                switch (lineText[j].location) {
                    case '2273,539,374,39':// bookingNumber
                        bookingNum = lineText[j].text.trim();
                        break;
                    case '135,685,404,49':// shipper(하주회사)
                        shipper = lineText[j].text.split('CO. ,LTD.')[0].trim();
                        break;
                    case '134,747,333,50':// shipper region
                        shipperRegion = lineText[j].text.trim();
                        break;
                    case '138,1069,995,51':// consignee(수취회사)
                        cosignee = lineText[j].text.split('CO.,LTD')[0].trim();
                        break;
                    case '138,1195,1244,53':// consignee region
                        cosigneeRegion = lineText[j].text.trim();
                        break;
                }
            }
        }
        else if (lineText[i].location == '1242,422,762,82' && lineText[i].text.trim() == 'SHIPPING REQUEST') {  // 05
            documentType = 2;
            for (var j = 0; j < lineText.length; j++) {
                switch (lineText[j].location) {
                    case '2248,616,383,38':// bookingNumber
                        bookingNum = lineText[j].text.trim();
                        break;
                    case '201,755,338,48':// shipper(하주회사)
                        shipper = lineText[j].text.split('00. ,LTD.')[0].trim();
                        break;
                    case '202,816,319,47':// shipper region
                        shipperRegion = lineText[j].text.trim();
                        break;
                    case '204,1248,873,46':// consignee(수취회사)
                        cosignee = lineText[j].text.split('CO. ,LTD.')[0].trim();
                        break;
                    case '203,1308,562,48':// consignee region
                        cosigneeRegion = lineText[j].text.trim();
                        break;
                }
            }
        }
        else if (lineText[i].location == '1236,420,773,84' && lineText[i].text.trim() == 'SHIPPING REQUEST') {  // 09, 11, 12, 13
            documentType = 2;
            bookingNum = '문서번호 없음';
            for (var j = 0; j < lineText.length; j++) {
                switch (lineText[j].location) {
                    //09
                    case '201,754,914,49':// shipper(하주회사)
                        shipper = lineText[j].text.split('00. ,LTD. ')[0].trim();
                        break;
                    case '201,815,336,49':// shipper region
                        shipperRegion = lineText[j].text.trim();
                        break;
                    case '203,1123,1063,49':// consignee(수취회사)
                        cosignee = lineText[j].text.split('CO. , LTD.')[0].trim();
                        break;
                    case '203,1247,1008,49':// consignee region
                        cosigneeRegion = lineText[j].text.trim();
                        break;

                    //11
                    case '201,754,717,50':// shipper(하주회사)
                        shipper = lineText[j].text.split('00. ,LTD. ')[0].trim();
                        break;

                    //12
                    case '202,753,996,50':// shipper(하주회사)
                        shipper = lineText[j].text.split('00. ,LTD. ')[0].trim();
                        break;

                    //13
                    case '201,751,1057,49':// shipper(하주회사)
                        shipper = lineText[j].text.trim();
                        break;
                }
            }
        }
        else if (lineText[i].location == '286,462,216,34' && lineText[i].text.trim() == 'SHIPPER:') {   // 04, 06, 07, 08, 10
            console.log('shipper');
            documentType = 2;
            for (var j = 0; j < lineText.length; j++) {
                switch (lineText[j].location) {
                    //04
                    case '1694,468,808,47':// bookingNumber
                        bookingNum = lineText[j].text.split('Booking NO. :')[1].trim();
                        break;
                    case '316,517,706,43':// shipper(하주회사)
                        shipper = lineText[j].text.split('CO. ,LTD.')[0].trim();
                        break;
                    case '316,638,547,42':// shipper region
                        shipperRegion = lineText[j].text.trim();
                        break;
                    case '316,894,370,42':// consignee(수취회사)
                        cosignee = lineText[j].text.split('CO., LTD.')[0].trim();
                        break;
                    case '316,1013,593,39':// consignee region
                        cosigneeRegion = lineText[j].text.trim();
                        break;

                    //06, 08, 10
                    case '1693,467,810,49':// bookingNumber
                        bookingNum = lineText[j].text.split('Booking NO. :')[1].trim();
                        break;
                    case '315,517,707,45':// shipper(하주회사)
                        shipper = lineText[j].text.split('00. , LTD.')[0].trim();
                        break;
                    case '315,637,549,45':// shipper region
                        shipperRegion = lineText[j].text.trim();
                        break;
                    case '315,1270,373,44':// consignee(수취회사)
                        cosignee = lineText[j].text.split('00. , LTD.')[0].trim();
                        break;
                    case '315,1390,595,38':// consignee region
                        cosigneeRegion = lineText[j].text.trim();
                        break;

                    //07, 10
                    case '315,577,779,44':// shipper(하주회사)
                        shipper = lineText[j].text.split('00. ,LTD.')[0].trim();
                        break;
                    case '315,637,764,45':// shipper region
                        shipperRegion = lineText[j].text.trim();
                        break;
                    case '315,893,452,38':// consignee + consignee region 
                        cosignee = lineText[j].text.trim();
                        cosigneeRegion = lineText[j].text.trim();
                        break;

                    //10
                    case '1693,467,809,49':// bookingNumber
                        bookingNum = lineText[j].text.split('Booking NO. :')[1].trim();
                        break;
                }
            }
        }
    }
    fileName = validationCheck(fileName);
    docDate = validationCheck(docDate);
    companyName = validationCheck(companyName);
    contractName = validationCheck(contractName);
    email = validationCheck(email);
    price1 = validationCheck(price1);
    price2 = validationCheck(price2);
    price3 = validationCheck(price3);
    totPrice = validationCheck(totPrice);
    detail = validationCheck(detail);
    etc = validationCheck(etc);

    send_company = validationCheck(send_company);
    send_tel = validationCheck(send_tel);
    send_fax = validationCheck(send_fax);
    send_person = validationCheck(send_person);
    receive_company = validationCheck(receive_company);
    receive_tel = validationCheck(receive_tel);
    receive_fax = validationCheck(receive_fax);
    receive_person = validationCheck(receive_person);

    bookingNum = validationCheck(bookingNum);
    shipper = validationCheck(shipper);
    shipperRegion = validationCheck(shipperRegion);
    cosignee = validationCheck(cosignee);
    cosigneeRegion = validationCheck(cosigneeRegion);

    excelParams.push({
        'fileName': fileName,
        'docDate': docDate,
        'companyName': companyName,
        'contractName': contractName,
        'email': email,
        'price1': price1,
        'price2': price2,
        'price3': price3,
        'totPrice': totPrice,
        'detail': detail,
        'etc': etc,
        
        'bookingNum': bookingNum,
        'shipper': shipper,
        'shipperRegion': shipperRegion,
        'cosignee': cosignee,
        'cosigneeRegion': cosigneeRegion,

        'documentType': documentType
    });

    if (documentType==1 && contractName == '' &&
        email == '' && price1 == '' && price2 == '' &&
        price3 == '' && totPrice == '') {
        $('.t04').text(Number($('.t04').text()) + 1);
    }
    else if (documentType == 2 && shipper == '' && cosignee == ''){
        $('.t04').text(Number($('.t04').text()) + 1);
    }
    else {
        $('.t01').text(Number($('.t01').text()) + 1);
    }

    if (checkNum == (Number($('.t01').text()) + Number($('.t04').text()))) {
        $('.tex01').text('완료');
        checkNum = 0;
    }
}

function validationCheck(value) {
    if (value == null || value == undefined || value.trim() == '') {
        return '';
    } else {
        return value.trim();
    }
}

function excelSaveEvent() {
    $('#sucessExcelBtn').click(function () {
        if ($('.t01').text() == '0') {
            alert('확인된 양식의 문서가 없습니다.');
        } else {
            analysisImg('success');
            //window.location = '/downloadExcel?fileName=' + $('#successExcelName').val();
        }
    });
    $('#failExcelBtn').click(function () {
        if ($('.t04').text() == '0') {
            alert('미확인된 양식의 문서가 없습니다.');
        } else {
            analysisImg('fail');
            //window.location = '/downloadExcel?fileName=' + $('#failExcelName').val();
        }
    });
}