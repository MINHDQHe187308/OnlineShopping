
let currentPrice = 0;
let currentVariantId = 0;
let productName = "";
let mainImageUrl = "";


function changeMainImage(thumb) {
    const mainImg = document.getElementById('mainImg');
    if (!mainImg) return;

    mainImg.src = thumb.dataset.full;
    document.querySelectorAll('.thumbnail').forEach(t => t.classList.remove('active'));
    thumb.classList.add('active');
}


function changeQty(delta) {
    const input = document.getElementById('quantity');
    if (!input) return;

    let qty = parseInt(input.value) || 1;
    qty = Math.max(1, qty + delta);
    input.value = qty;
    updateTotalPrice();
}

function showCartMessage(message, isSuccess) {
    const messageEl = document.getElementById('cart-message');
    if (!messageEl) {
        alert(message);
        return;
    }

    messageEl.textContent = message;
    messageEl.classList.remove('d-none', 'alert-success', 'alert-danger');
    messageEl.classList.add(isSuccess ? 'alert-success' : 'alert-danger');

    setTimeout(() => {
        messageEl.classList.add('d-none');
    }, 3000);
}

function updateTotalPrice() {
    const input = document.getElementById('quantity');
    const totalEl = document.getElementById('totalPrice');
    if (!input || !totalEl) return;

    const qty = parseInt(input.value) || 1;
    const total = currentPrice * qty;
    totalEl.textContent = total.toLocaleString('vi-VN') + ' đ';
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
    } catch (err) {
        console.error('Lỗi cập nhật số giỏ hàng:', err);
    }
}


function setupListAddToCart() {
    document.addEventListener('click', async function (e) {
        const btn = e.target.closest('.add-to-cart-btn');
        if (!btn) return;

        e.preventDefault();

        const variantId = parseInt(btn.dataset.variantId);
        const qty = 1; 
        const price = parseFloat(btn.dataset.price) || 0;
        const name = btn.dataset.productName || "Sản phẩm";
        const img = btn.dataset.image || "/images/no-image.jpg";

        if (variantId <= 0) {
            alert('Sản phẩm này hiện không có biến thể hợp lệ!');
            return;
        }

        try {
            const tokenElement = document.querySelector('input[name="__RequestVerificationToken"]');
            const token = tokenElement ? tokenElement.value : null;

            const response = await fetch('/Cart/AddToCart', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    ...(token && { 'RequestVerificationToken': token })
                },
                body: JSON.stringify({ variantId: variantId, quantity: qty })
            });

            if (response.status === 401) {
                window.location.href = '/Login';
                return;
            }

            if (!response.ok) {
                const errText = await response.text();
                throw new Error(errText || 'Có lỗi xảy ra khi thêm vào giỏ');
            }

            // Cập nhật badge
            await updateCartCount();

            // Hiển thị modal thành công
            const modalEl = document.getElementById('addToCartSuccessModal');
            if (modalEl) {
                document.getElementById('modalProductImage').src = img;
                document.getElementById('modalProductName').textContent = name;
                document.getElementById('modalSize').textContent = "Mặc định"; // Có thể cải thiện nếu có size
                document.getElementById('modalQuantity').textContent = qty;
                document.getElementById('modalPrice').textContent = (price * qty).toLocaleString('vi-VN');

                const modal = new bootstrap.Modal(modalEl);
                modal.show();
            } else {
                alert(`Đã thêm "${name}" vào giỏ hàng!`);
            }

        } catch (err) {
            alert('Không thể thêm sản phẩm: ' + err.message);
            console.error(err);
        }
    });
}

document.addEventListener('DOMContentLoaded', () => {
    // 1. Khởi tạo cho trang CHI TIẾT sản phẩm
    const dataEl = document.getElementById('product-data');
    if (dataEl) {
        currentPrice = parseFloat(dataEl.dataset.price) || 0;
        currentVariantId = parseInt(dataEl.dataset.variantId) || 0;
        productName = dataEl.dataset.name || "";
        mainImageUrl = dataEl.dataset.imageUrl || "/images/no-image.jpg";

        updateTotalPrice();
        const qtyInput = document.getElementById('quantity');
        if (qtyInput) {
            qtyInput.addEventListener('input', () => {
                if (parseInt(qtyInput.value) < 1 || isNaN(parseInt(qtyInput.value))) {
                    qtyInput.value = 1;
                }
                updateTotalPrice();
            });
        }

        // Sự kiện chọn size (trang chi tiết)
        document.querySelectorAll('.size-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                document.querySelectorAll('.size-btn').forEach(b => b.classList.remove('active'));
                btn.classList.add('active');

                currentPrice = parseFloat(btn.dataset.price) || 0;
                currentVariantId = parseInt(btn.dataset.variantId) || 0;

                const priceEl = document.getElementById('currentPrice');
                if (priceEl) {
                    priceEl.textContent = currentPrice.toLocaleString('vi-VN') + ' đ';
                }

                updateTotalPrice();
            });
        });

       
        //const addBtn = document.getElementById('addToCart');
        //if (addBtn) {
        //    addBtn.addEventListener('click', async (e) => {
        //        e.preventDefault();

        //        const qtyInput = document.getElementById('quantity');
        //        const qty = qtyInput ? parseInt(qtyInput.value) || 1 : 1;

        //        const activeSize = document.querySelector('.size-btn.active');
        //        const sizeText = activeSize ? activeSize.textContent.trim() : 'Không chọn';

        //        if (currentVariantId <= 0) {
        //            alert('Vui lòng chọn kích thước!');
        //            return;
        //        }

        //        try {
        //            const token = document.querySelector('input[name="__RequestVerificationToken"]')?.value;

        //            const response = await fetch('/Cart/AddToCart', {
        //                method: 'POST',
        //                headers: {
        //                    'Content-Type': 'application/json',
        //                    ...(token && { 'RequestVerificationToken': token })
        //                },
        //                body: JSON.stringify({ variantId: currentVariantId, quantity: qty })
        //            });

        //            if (!response.ok) {
        //                const errText = await response.text();
        //                throw new Error(errText || 'Có lỗi xảy ra');
        //            }

        //            // Modal thành công
        //            const modalEl = document.getElementById('addToCartSuccessModal');
        //            if (modalEl) {
        //                document.getElementById('modalProductImage').src = mainImageUrl;
        //                document.getElementById('modalProductName').textContent = productName;
        //                document.getElementById('modalSize').textContent = sizeText;
        //                document.getElementById('modalQuantity').textContent = qty;
        //                document.getElementById('modalPrice').textContent = (currentPrice * qty).toLocaleString('vi-VN');

        //                const modal = new bootstrap.Modal(modalEl);
        //                modal.show();
        //            }

        //            updateCartCount();

        //        } catch (err) {
        //            alert('Không thể thêm vào giỏ: ' + err.message);
        //            console.error(err);
        //        }
        //    });
        //}
        const addBtn = document.getElementById('addToCart');
        if (addBtn) {
            addBtn.addEventListener('click', async (e) => {
                e.preventDefault();

                const qtyInput = document.getElementById('quantity');
                const qty = qtyInput ? parseInt(qtyInput.value) || 1 : 1;

                const activeSize = document.querySelector('.size-btn.active');
                const sizeText = activeSize ? activeSize.textContent.trim() : 'Không chọn';

                if (currentVariantId <= 0) {
                    showCartMessage('Vui lòng chọn kích thước!', false);
                    return;
                }

                try {
                    const token = document.querySelector('input[name="__RequestVerificationToken"]')?.value;

                    const response = await fetch('/Cart/AddToCart', {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/json',
                            ...(token && { 'RequestVerificationToken': token })
                        },
                        body: JSON.stringify({ variantId: currentVariantId, quantity: qty })
                    });

                    if (response.status === 401) {
                        window.location.href = '/Login';
                        return;
                    }

                    const result = await response.json().catch(() => null);

                    if (!response.ok) {
                        showCartMessage(result?.message || 'Không thể thêm vào giỏ hàng!', false);
                        return;
                    }

                    await updateCartCount();
                    showCartMessage(result?.message || 'Đã thêm vào giỏ hàng!', true);

                } catch (err) {
                    showCartMessage('Không thể thêm vào giỏ hàng!', false);
                    console.error(err);
                }
            });
        }
    }

    setupListAddToCart();

  
    updateCartCount();
    document.querySelectorAll(".change-image-input").forEach(input => {

        input.addEventListener("change", function () {

            const file = this.files[0];
            const productId = this.dataset.productId;

            if (!file) return;

            const formData = new FormData();
            formData.append("productId", productId);
            formData.append("imageFile", file);

            fetch("/Product/UpdateProductImage", {
                method: "POST",
                body: formData
            })
                .then(res => res.json())
                .then(data => {

                    if (data.success) {

                        const card = this.closest(".product-card");
                        const img = card.querySelector("img");

                        img.src = data.newImageUrl + "?t=" + new Date().getTime();
                    }
                });

        });

    });
});