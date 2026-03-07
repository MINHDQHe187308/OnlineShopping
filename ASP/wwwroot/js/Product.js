let currentPrice = @Model.CurrentPrice;

// Change main image
function changeMainImage(thumb) {
    document.getElementById('mainImg').src = thumb.dataset.full;
    document.querySelectorAll('.thumbnail').forEach(t => t.classList.remove('active'));
    thumb.classList.add('active');
}

// Change quantity & update total
function changeQty(delta) {
    const qtyInput = document.getElementById('quantity');
    let qty = parseInt(qtyInput.value) || 1;
    qty = Math.max(1, qty + delta);
    qtyInput.value = qty;
    updateTotalPrice();
}

// Update total price
function updateTotalPrice() {
    const qty = parseInt(document.getElementById('quantity').value) || 1;
    const total = currentPrice * qty;
    document.getElementById('totalPrice').textContent = total.toLocaleString('vi-VN') + ' đ';
}

// Size selection - update price
document.querySelectorAll('.size-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        document.querySelectorAll('.size-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');

        currentPrice = parseFloat(btn.dataset.price) || 0;
        document.getElementById('currentPrice').textContent = currentPrice.toLocaleString('vi-VN') + ' đ';
        updateTotalPrice();
    });
});

// Add to cart (placeholder - sau này thay bằng AJAX)
document.getElementById('addToCart').addEventListener('click', () => {
    const selectedSize = document.querySelector('.size-btn.active')?.dataset.size || 'Không chọn';
    const qty = document.getElementById('quantity').value;

    alert(`Đã thêm vào giỏ:\nSản phẩm: ${@Model.Product.ProductName}\nSize: ${selectedSize}\nSố lượng: ${qty}\nGiá: ${currentPrice.toLocaleString('vi-VN')} đ`);
    // Thực tế: gửi fetch POST đến /Cart/Add
});