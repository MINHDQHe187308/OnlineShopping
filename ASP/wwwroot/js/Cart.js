
function updateCart() {

    let total = 0;

    document.querySelectorAll("tbody tr").forEach(row => {

        let price = parseFloat(row.dataset.price);

        let qtyInput = row.querySelector(".quantity");

        let qty = parseInt(qtyInput.value);

        let subtotal = price * qty;

        row.querySelector(".subtotal").innerText =
            subtotal.toLocaleString() + " đ";

        total += subtotal;

    });

    document.getElementById("cartTotal").innerText =
        total.toLocaleString() + " đ";

}

document.querySelectorAll(".plus").forEach(btn => {

    btn.onclick = function () {

        let input = this.parentElement.querySelector(".quantity");

        input.value = parseInt(input.value) + 1;

        updateCart();

    }

});

document.querySelectorAll(".minus").forEach(btn => {

    btn.onclick = function () {

        let input = this.parentElement.querySelector(".quantity");

        if (input.value > 1) {

            input.value = parseInt(input.value) - 1;

            updateCart();

        }

    }

});

document.querySelectorAll(".quantity").forEach(input => {

    input.onchange = function () {

        if (this.value < 1) this.value = 1;

        updateCart();

    }

});