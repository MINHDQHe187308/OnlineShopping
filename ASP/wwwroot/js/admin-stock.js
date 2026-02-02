$(document).ready(function () {
    // dong bo cigma
    $("#btn-synchronize").click(function (e) {
        e.preventDefault();
        let confirmMsg = "Bạn muốn thực hiện tác vụ này?";
        if (confirm(confirmMsg) == true) {
            let txtStockYear = $("#txtStockYear").val();
            if (txtStockYear == "" || txtStockYear == null || txtStockYear == "undefined") return alert("Năm không để trống.");
            let txtStockMonth = $("#txtStockMonth").val();
            if (txtStockMonth == "" || txtStockMonth == null || txtStockMonth == "undefined") return alert("Tháng không để trống.");
            //
            let txtHouse = $("#txtHouse").val();
            if (txtHouse == "" || txtHouse == null || txtHouse == "undefined") return alert("Kho không để trống.");
            let txtPartNo = $("#txtPartNo").val();
            let txtRecord = $("#txtRecord").val();
            if (txtRecord == "" || txtRecord == null || txtRecord == "undefined") return alert("Số bản ghi không để trống.");
            //Set the URL.
            $(this).prop("disabled", true);
            let url = $("#frm-synchronize").attr("action");
            //Add the Field values to FormData object.
            let formData = new FormData();
            formData.append("year", txtStockYear);
            formData.append("month", txtStockMonth);
            formData.append("house", txtHouse);
            formData.append("partno", txtPartNo);
            formData.append("nRecord", txtRecord);
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
    
    // jQuery
});