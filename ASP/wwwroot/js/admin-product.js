$(document).ready(function () {
    // dong bo cigma
    $("#btn-synchronize").click(function (e) {
        e.preventDefault();
        let confirmMsg = "Bạn muốn thực hiện tác vụ này?";
        if (confirm(confirmMsg) == true) {
            let txtHouse = $("#txtHouse").val();
            if (txtHouse == "" || txtHouse == null || txtHouse == "undefined") return alert("Kho không để trống.");
            let txtPartNo = $("#txtPartNo").val();
            let txtRecord = $("#txtRecord").val();
            if (txtRecord == "" || txtRecord == null || txtRecord == "undefined") return alert("Số bản ghi không để trống.");
            //
            var chkRiskStock = $('#frmchkRiskStock').is(":checked");
            var fexRiskStock = false;
            if (chkRiskStock) {
                // checked
                fexRiskStock = true;
            }

            var chkStatusNCC = $('#frmchkStatusNCC').is(":checked");
            var fexStatusNCC = false;
            if (chkStatusNCC) {
                // checked
                fexStatusNCC = true;
            }

            var chkStatus = $('#frmchkStatus').is(":checked");
            var fexStatus = false;
            if (chkStatus) {
                // checked
                fexStatus = true;
            }
            //Set the URL.
            $(this).prop("disabled", true);
            let url = $("#frm-synchronize").attr("action");
            //Add the Field values to FormData object.
            let formData = new FormData();
            formData.append("house", txtHouse);
            formData.append("partno", txtPartNo);
            formData.append("nRecord", txtRecord);
            formData.append("nRiskStock", fexRiskStock);
            formData.append("nStatusNCC", fexStatusNCC);
            formData.append("nStatus", fexStatus);
            //
            $.ajax({
                type: "POST",
                url: url,
                headers: { "RequestVerificationToken": $("form").find("input[name=__RequestVerificationToken]").val() },
                data: formData,
                processData: false,
                contentType: false,
                beforeSend: function () {
                    $('#sct-loading').show();
                    $("#btn-synchronize").removeAttr('disabled');
                },
                success: function (data) {
                    $("#btn-synchronize").removeAttr('disabled');
                    $('#sct-loading').hide();
                    $('.box-message-synchronize').html(data.message);
                    $(".bd-synchronize-modal-lg").on('hide.bs.modal', function () {
                        location.reload();
                    });
                },
                error: function (error) {
                    $("#btn-synchronize").removeAttr('disabled');
                    $('#sct-loading').hide();
                    alert("Có lỗi xảy ra khi thực hiện tác vụ.");
                }
            });
        } else {
            console.log("You canceled!");
        }
        //
    });
    // import master data
    // event upload btn
    $('#btn-import').click(function (e) {
        e.preventDefault();
        let importFile = $("#txtfile").val();
        if (!importFile) {
            alert("Bạn cần chọn đường dẫn của file cần import.");
            return false;
        }
        let confirmMsg = "Bạn có chắc muốn import những yêu cầu tuyển chọn từ file này?";
        if (confirm(confirmMsg) == true) {
            // use ajax call back action
            $(this).prop("disabled", true);
            let url = $("#frmImport").attr("action");
            let formData = new FormData($('form#frmImport')[0]);
            $.ajax({
                type: 'POST',
                url: url,
                headers: { "RequestVerificationToken": $("form#frmImport").find("input[name=__RequestVerificationToken]").val() },
                data: formData,
                cache: false,
                contentType: false,
                processData: false,
                beforeSend: function () {
                    $('#sct-loading').show();
                    $("#btn-import").removeAttr('disabled');
                },
                success: function (data) {
                    $("#btn-import").removeAttr('disabled');
                    $('#sct-loading').hide();
                    $('.box-message-import').html(data.message);
                    $(".bd-import-modal-lg").on('hide.bs.modal', function () {
                        location.reload();
                    });
                },
                error: function () {
                    $("#btn-import").removeAttr('disabled');
                    $('#sct-loading').hide();
                    alert("error");
                }
            });
        } else {
            console.log("You canceled!");
        }
        //
    });
    // jQuery
});