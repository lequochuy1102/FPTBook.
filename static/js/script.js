// const toggler = document.querySelector(".btn");
// toggler.addEventListener("click",function(){
//     document.querySelector("#sidebar").classList.toggle("collapsed");
// });

const toggler = document.querySelector(".btn");
const sidebar = document.querySelector("#sidebar");

// Đảm bảo sidebar thu lại khi trang được tải
document.addEventListener("DOMContentLoaded", function () {
    sidebar.classList.add("collapsed");
});

// Sự kiện bấm nút để toggle sidebar
toggler.addEventListener("click", function (event) {
    event.stopPropagation(); // Ngăn sự kiện click lan ra ngoài
    sidebar.classList.toggle("collapsed");
});

// Sự kiện bấm bất kỳ chỗ nào trên tài liệu
document.addEventListener("click", function (event) {
    // Kiểm tra nếu click không nằm trong sidebar hoặc nút toggle
    if (!sidebar.contains(event.target) && !toggler.contains(event.target)) {
        sidebar.classList.add("collapsed");
    }
});

