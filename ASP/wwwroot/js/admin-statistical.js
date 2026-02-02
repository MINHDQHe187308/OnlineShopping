$(document).ready(function () {
    // import master data
    // event upload btn
    $('#btn-import').click(function (e) {
        e.preventDefault();
        let impYear = $("#impYear").val();
        if (impYear == "" || impYear == null) {
            alert("Năm import dữ liệu không để trống.");
            return false;
        }
        let impCat = $("#impCat").val();
        if (impCat == "" || impCat == null) {
            alert("Loại hàng không để trống.");
            return false;
        }
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