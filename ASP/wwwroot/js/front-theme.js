$(document).ready(function () {
    // pagination
    $('.slt-psize').change(function () {
        var strVal = $(this).val();
        $('#psize').val(strVal);
        $('form.search-form').submit()
    });
});
function getLoading() {
    var str = '';
    str += '<p class="p-loading"><span><i class="fas fa-spinner fa-pulse"></i></span></p>'
    $('body').html(str);
}
function removeLoading() {
    $('p.p-loading').remove();
}
function setHeigthBody() {
    var w = $(window).height();
    var header = $("#header").height();
    var footer = $("#footer").height();
    $("div#wrap-content").css("min-height", (w - header - footer) + "px");
}