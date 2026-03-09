
function updateCartTotal() {
    let total = 0;
    document.querySelectorAll("tbody tr").forEach(row => {
        const price = parseFloat(row.dataset.price) || 0;
        const qtyInput = row.querySelector(".quantity");
        const qty = parseInt(qtyInput?.value) || 1;
        const subtotal = price * qty;

        const subtotalEl = row.querySelector(".subtotal");
        if (subtotalEl) {
            subtotalEl.innerText = subtotal.toLocaleString('vi-VN') + " đ";
        }

        total += subtotal;
    });

    
    const cartTotalEl = document.getElementById("cartTotal");
    if (cartTotalEl) {
        cartTotalEl.innerText = total.toLocaleString('vi-VN') + " đ";
    }
}


async function updateCartCount() {
    try {
        const res = await fetch('/Cart/GetCartCount');
        if (!res.ok) return;
        const count = await res.json();

        const badges = document.querySelectorAll('.cart-badge, .cart-count-badge, #cartCount, .badge-cart');
        badges.forEach(badge => {
            badge.textContent = count;
        });

        
        if (count === 0) {
            setTimeout(() => location.reload(), 500);
        }
    } catch (err) {
        console.error('Lỗi cập nhật số giỏ:', err);
    }
}


function setupQuantityButtons() {
    document.querySelectorAll(".plus").forEach(btn => {
        btn.onclick = function () {
            const input = this.parentElement.querySelector(".quantity");
            input.value = parseInt(input.value) + 1;
            updateCartTotal();
          
        };
    });

    document.querySelectorAll(".minus").forEach(btn => {
        btn.onclick = function () {
            const input = this.parentElement.querySelector(".quantity");
            if (input.value > 1) {
                input.value = parseInt(input.value) - 1;
                updateCartTotal();
            }
        };
    });

    document.querySelectorAll(".quantity").forEach(input => {
        input.onchange = function () {
            if (this.value < 1) this.value = 1;
            updateCartTotal();
        };
    });
}


function setupRemoveItem() {
    document.addEventListener('click', async function (e) {
        const btn = e.target.closest('.remove-item');
        if (!btn) return;

        e.preventDefault();

        const cartItemId = parseInt(btn.dataset.cartitemId);
        if (!cartItemId || isNaN(cartItemId)) {
            alert('Không tìm thấy thông tin sản phẩm để xóa!');
            return;
        }

        if (!confirm('Bạn có chắc muốn xóa sản phẩm này khỏi giỏ hàng?')) {
            return;
        }

        try {
            const token = document.querySelector('input[name="__RequestVerificationToken"]')?.value;

            const response = await fetch('/Cart/RemoveItem', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    ...(token && { 'RequestVerificationToken': token })
                },
                body: JSON.stringify({ cartItemId: cartItemId })
            });

            if (!response.ok) {
                const errText = await response.text();
                throw new Error(errText || 'Xóa thất bại');
            }

      
            const row = btn.closest('tr');
            if (row) {
                row.remove();
            }

            updateCartTotal();
            await updateCartCount();

          

        } catch (err) {
            alert('Không thể xóa sản phẩm: ' + err.message);
            console.error(err);
        }
    });
}

document.addEventListener('DOMContentLoaded', () => {
   
    updateCartTotal();          
    setupQuantityButtons();   
    setupRemoveItem();        
});