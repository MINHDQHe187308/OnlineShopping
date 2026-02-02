const Toast = Swal.mixin({
    toast: true,
    position: 'top-end',
    showConfirmButton: false,
    timer: 5000
});
function msgErr(msg) {
    Toast.fire({ icon: "warning", title: msg });
}

function msgSuc(msg) {
    Toast.fire({ icon: "success", title: msg });
}

function cancelTask() {
    msgSuc("Bạn đã hủy tác vụ!");
}
function isNullOrEmpty(str) {
    if (str === null) return true;
    if (str === "") return true;
    if (str.length == 0) return true;
    return false;
}