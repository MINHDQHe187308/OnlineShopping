$(document).ready(function () {
    // pagination
    $('.slt-psize').change(function () {
        var strVal = $(this).val();
        $('#psize').val(strVal);
        $('form.search-form').submit()
    });
    // load url image
    $(".image-input").change(function () {
        var element = $('.image-preview');
        readURL(element, this);
        //
        $('.rmavatar').val('');
    });
    // remove avatar
    $('.preview-image span.iconx-preview').click(function () {
        $('.rmavatar').val('remove');
        $('.image-preview').attr('src', '');
        //
        $('.image-input-user').val('');
        //
        $('.pdf-preview').attr('src', '');
    });
    //
    /** 
     * load url image 2
     * */
    $(".image-input2").change(function () {
        var element = $('.image-preview2');
        readURL(element, this);
        //
        $('.rmavatar2').val('');
    });
    // remove avatar
    $('.preview-image2 span.iconx-preview2').click(function () {
        $('.rmavatar2').val('remove');
        $('.image-preview2').attr('src', '');
        //
        $('.image-input-user2').val('');
        //
    });

    // end document ready;
});
function readURL(element, input) {

    if (input.files && input.files[0]) {
        if (input.files[0].size > 2000000) {
            $(input).next().removeClass('hidden');
            $(input).next().html('Chá»‰ upload áº£nh < 2M')
        } else {
            $(input).next().html('');
            $(input).next().addClass('hidden');
            var reader = new FileReader();
            reader.onload = function (e) {
                element.attr('src', e.target.result);
            };
            reader.readAsDataURL(input.files[0]);
        }
    }
}
//
