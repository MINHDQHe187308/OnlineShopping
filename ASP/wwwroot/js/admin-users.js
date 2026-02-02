
$(document).ready(function () {
    /**
     * Select Province
     * @type {Choices}
     */
    const choices_province = new Choices('.tags-input-province', {
        delimiter: ',',
        removeItemButton: true,
        placeholder: true,
    });

    choices_province.passedElement.addEventListener('addItem', function (event) {
        let array = JSON.parse(event.detail.value);
        let provinceElement = $('.provinces');
        let provinces = provinceElement.val();
        provinces = JSON.parse(provinces ? provinces : '[]');
        provinces.push({province_id: array[0] });
        provinceElement.val(JSON.stringify(provinces));
    }, false);
    choices_province.passedElement.addEventListener('removeItem', function () {
        let array = choices_province.getValue();
        let provinceElement = $('.provinces');
        let provinces = [];
        array.forEach(function (value) {
            let convert = JSON.parse(value.value);
            provinces.push({ province_id: convert[0] });
        });
        provinceElement.val(JSON.stringify(provinces));
    }, false);

    // end document ready;
});



