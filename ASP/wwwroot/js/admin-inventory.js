$(document).ready(function () {
    //
    $('#reservationdate').datetimepicker({
        singleDatePicker: true,
        format: 'MM/DD/YYYY'
    });
    // dong bo cigma
    $("#btn-synchronize").click(function (e) {
        e.preventDefault();
        let confirmMsg = "Bạn muốn thực hiện tác vụ này?";
        if (confirm(confirmMsg) == true) {
            let txtProductDate = $("#txtProductDate").val();
            if (txtProductDate == "" || txtProductDate == null || txtProductDate == "undefined") return alert("Ngày sản xuất không để trống.");
            //
            let txtHouse = $("#txtHouse").val();
            if (txtHouse == "" || txtHouse == null || txtHouse == "undefined") return alert("Kho không để trống.");
            let txtPartNo = $("#txtPartNo").val();
            //let txtRecord = $("#txtRecord").val();
            //if (txtRecord == "" || txtRecord == null || txtRecord == "undefined") return alert("Số bản ghi không để trống.");
            //Set the URL.
            $(this).prop("disabled", true);
            let url = $("#frm-synchronize").attr("action");
            //Add the Field values to FormData object.
            let formData = new FormData();
            formData.append("productDate", txtProductDate);
            formData.append("house", txtHouse);
            formData.append("partno", txtPartNo);
            //formData.append("nRecord", txtRecord);
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
                    if (data.status == "error") {
                        $('.box-message-synchronize').addClass("bg-danger");
                        $('.box-message-synchronize').removeClass("bg-warning");
                    } else {
                        $('.box-message-synchronize').addClass("bg-warning");
                        $('.box-message-synchronize').removeClass("bg-danger");
                    }
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