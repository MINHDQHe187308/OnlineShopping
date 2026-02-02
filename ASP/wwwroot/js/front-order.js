$(document).ready(function () {
    $('#btn-Change').click(function (e) {
        let _uId = document.getElementById('frm_change_uId').value;
        if (isNullOrEmpty(_uId)) {
            msgErr("Không xác nhận được Order cần cập nhật!");
            return false;
        }

        let cDate = document.getElementById('frm_change_cDate').value;
        if (cDate === null || cDate === "") {
            msgErr("Bạn cần chọn ngày giao hàng!");
            return false;
        }

        let confirmMsg = "Bạn muốn thực hiện tác vụ này?";

        let formData = new FormData($('form#frmChange')[0]);
        formData.append("uId", _uId);
        formData.append("cDate", cDate);

        if (confirm(confirmMsg) == true) {
            $.ajax({
                type: 'POST',
                url: '@Url.Action("ChangeShippingDate", "Order")',
                headers: { "RequestVerificationToken": $("form#frmChange").find("input[name=__RequestVerificationToken]").val() },
                data: formData,
                cache: false,
                contentType: false,
                processData: false,
                beforeSend: function () {
                    $('#sct-loading').show();
                    $("#btn-Change").removeAttr('disabled');
                },
                success: function (data) {
                    $("#btn-Change").removeAttr('disabled');
                    $('#sct-loading').hide();

                    if (data.result == "success") {
                        $('#frmModalChange').modal("hide");
                        window.location.reload();
                        msgSuc(data.message);
                    } else {
                        msgErr(data.message);
                    }
                },
                complete: function () {
                    $('#frmModalChange').modal("hide");
                    $('#sct-loading').hide();
                },
                error: function () {
                    $("#btn-Change").removeAttr('disabled');
                    $('#sct-loading').hide();
                    msgErr(data.message);
                }
            });
        } else {
            cancelTask();
        }
    });

    $('#btn-Download-Odr').click(function (e) {
        let formData = new FormData($('form#frmDownloadOdr')[0]);

        formData.append("uId", document.getElementById('frm_download_uid').value);
        formData.append("isOrder", getCheckBoxValue('frm_odr_check'));
        formData.append("isContent", getCheckBoxValue('frm_ctl_check'));
        formData.append("isKanban", getCheckBoxValue('frm_kbl_check'));

        $.ajax({
            type: 'POST',
            url: '@Url.Action("ProcessFile", "Order")',
            headers: { "RequestVerificationToken": $("form#frmDownloadOdr").find("input[name=__RequestVerificationToken]").val() },
            data: formData,
            cache: false,
            contentType: false,
            processData: false,
            beforeSend: function () {
                $('#sct-loading').show();
                $("#btn-Download-Odr").removeAttr('disabled');
            },
            success: function (data) {
                $("#btn-Download-Odr").removeAttr('disabled');
                $('#sct-loading').hide();

                if (data.result == true) {
                    $('#frmModalDownloadOdr').modal("hide");
                    debugger
                    window.location.href = "@Url.Action("DownloadFile", "Order")?uId=" + data.id + "&iM=" + data.im + "&iO=" + data.io + "&iC=" + data.ic + "&iK=" + data.ik;
                } else {
                    msgErr(data.message);
                }
            },
            complete: function () {
                $('#frmModalDownloadOdr').modal("hide");
                $('#sct-loading').hide();
            },
            error: function (data) {
                $("#btn-Download-Odr").removeAttr('disabled');
                $('#sct-loading').hide();
                msgErr(data.message);
            }
        });
    });
});

function openFrmChange(uId) {
    $('#frmModalChange').modal("show");
    $('#frm_change_uId').val(uId);
}

function openDownLoad(odrNo, uid) {
    $('#frmModalDownloadOdr').modal("show");
    document.getElementById('odrDAO').innerHTML = odrNo;
    $('#frm_download_uid').val(uid);
}

function actCancel(_uId) {
    let confirmMsg = "Bạn muốn thực hiện tác vụ này?";

    let formData = new FormData($('form#frmChange')[0]);
    formData.append("uId", _uId);

    if (confirm(confirmMsg) == true) {
        $.ajax({
            type: 'POST',
            url: '@Url.Action("CancelOrder", "Order")',
            headers: { "RequestVerificationToken": $("form#frmChange").find("input[name=__RequestVerificationToken]").val() },
            data: formData,
            cache: false,
            contentType: false,
            processData: false,
            beforeSend: function () {
                $('#sct-loading').show();
            },
            success: function (data) {
                $('#sct-loading').hide();
                if (data.result == "success") {
                    window.location.reload();
                    msgSuc(data.message);
                } else {
                    msgErr(data.message);
                }
            },
            complete: function () {
                $('#sct-loading').hide();
            },
            error: function () {
                $('#sct-loading').hide();
                msgErr(data.message);
            }
        });
    } else {
        cancelTask();
    }
}

function getCheckBoxValue(tagId) {
    let val = document.getElementById(tagId).checked;
    if (val) return "true";
    return "false";
}