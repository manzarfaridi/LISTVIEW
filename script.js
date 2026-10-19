
document.addEventListener("DOMContentLoaded", function () {
    const btnList = document.getElementById("btn-list-view");
    const btnGrid = document.getElementById("btn-grid-view");
    
    const listSection = document.getElementById("bda-list-view-section");
    const gridSection = document.getElementById("bda-grid-view-section");

    // 📝 लिस्ट व्यू पर क्लिक करने पर
    btnList.addEventListener("click", function () {
        btnList.classList.add("active");
        btnGrid.classList.remove("active");
        
        listSection.classList.replace("display-off", "display-on");
        gridSection.classList.replace("display-on", "display-off");
    });

    // 📊 ग्रिड व्यू पर क्लिक करने पर
    btnGrid.addEventListener("click", function () {
        btnGrid.classList.add("active");
        btnList.classList.remove("active");
        
        gridSection.classList.replace("display-off", "display-on");
        listSection.classList.replace("display-on", "display-off");
    });
});
//=================================================================//
//                     GRID VIEW SCRIPT STARTS
//=================================================================//

document.addEventListener("DOMContentLoaded", function () {
    const searchInput = document.querySelector(".bda-input-search");
    const allCompaniesCheckbox = document.getElementById("company-all");
    const singleCompaniesCheckboxes = document.querySelectorAll(".company-single");
    const allCategoriesCheckbox = document.getElementById("category-all");
    const singleCategoriesCheckboxes = document.querySelectorAll(".category-single");
    const productCards = document.querySelectorAll(".bda-product-card");
    const btnExcel = document.getElementById("btn-download-excel");
    const btnWhatsapp = document.getElementById("btn-share-whatsapp");

    allCompaniesCheckbox.addEventListener("change", function() {
        if (this.checked) singleCompaniesCheckboxes.forEach(cb => cb.checked = false);
        filterProducts();
    });
    singleCompaniesCheckboxes.forEach(cb => {
        cb.addEventListener("change", function() {
            if (this.checked) allCompaniesCheckbox.checked = false;
            if (!Array.from(singleCompaniesCheckboxes).some(c => c.checked)) allCompaniesCheckbox.checked = true;
            filterProducts();
        });
    });

    allCategoriesCheckbox.addEventListener("change", function() {
        if (this.checked) singleCategoriesCheckboxes.forEach(cb => cb.checked = false);
        filterProducts();
    });
    singleCategoriesCheckboxes.forEach(cb => {
        cb.addEventListener("change", function() {
            if (this.checked) allCategoriesCheckbox.checked = false;
            if (!Array.from(singleCategoriesCheckboxes).some(c => c.checked)) allCategoriesCheckbox.checked = true;
            filterProducts();
        });
    });

    function filterProducts() {
        const searchText = searchInput.value.toLowerCase().trim();
        const isAllCompanySelected = allCompaniesCheckbox.checked;
        const activeCompanies = [];
        singleCompaniesCheckboxes.forEach(cb => { if (cb.checked) activeCompanies.push(cb.value); });

        const isAllCategorySelected = allCategoriesCheckbox.checked;
        const activeCategories = [];
        singleCategoriesCheckboxes.forEach(cb => { if (cb.checked) activeCategories.push(cb.value); });

        productCards.forEach(card => {
            const brandName = card.getAttribute("data-brand").toLowerCase();
            const cardCompany = card.getAttribute("data-company");
            const cardCategory = card.getAttribute("data-category");

            const matchesSearch = brandName.includes(searchText);
            const matchesCompany = isAllCompanySelected || activeCompanies.includes(cardCompany);
            const matchesCategory = isAllCategorySelected || activeCategories.includes(cardCategory);

            if (matchesSearch && matchesCompany && matchesCategory) {
                card.style.display = "flex";
            } else {
                card.style.display = "none";
            }
        });
    }
    searchInput.addEventListener("input", filterProducts);

    function getSelectedProductsData() {
        const selectedData = [];
        productCards.forEach(card => {
            const isChecked = card.querySelector(".bda-product-select-check").checked;
            if (isChecked && card.style.display !== "none") {
                selectedData.push({
                    Brand: card.getAttribute("data-brand"),
                    Company: card.getAttribute("data-company").toUpperCase(),
                    Category: card.querySelector(".bda-cat-val").innerText,
                    Needle_Profile: card.querySelector(".bda-needle-val").innerText,
                    Suture_Profile: card.querySelector(".bda-suture-val").innerText,
                    MRP: card.querySelector(".bda-mrp-val").innerText
                });
            }
        });
        return selectedData;
    }

    btnExcel.addEventListener("click", function() {
        const data = getSelectedProductsData();
        if (data.length === 0) {
            alert("कृपया एक्सपोर्ट करने के लिए पहले 'चुनें (Select)' चेकबॉक्स पर टिक करें!");
            return;
        }
        const worksheet = XLSX.utils.json_to_sheet(data);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Selected Products");
        XLSX.writeFile(workbook, "Product_Matrix_Report.xlsx");
    });

    btnWhatsapp.addEventListener("click", function() {
        const data = getSelectedProductsData();
        if (data.length === 0) {
            alert("कृपया व्हाट्सएप पर भेजने के लिए पहले किसी उत्पाद को सेलेक्ट करें!");
            return;
        }
        let messageText = "*📦 Selected Products Report *\n\n";
        data.forEach((prod, index) => {
            messageText += `${index + 1}. *${prod.Brand}* (${prod.Company})\n`;
            messageText += `   • Cat: ${prod.Category}\n`;
            messageText += `   • Price: ${prod.MRP}\n\n`;
        });
        const whatsappUrl = `https://whatsapp.com{encodeURIComponent(messageText)}`;
        window.open(whatsappUrl, "_blank");
    });
});
//=================================================================//
//                      GRID VIEW SCRIPT STARTS
//=================================================================//
//===========================================================================
//        BIHAR DRUGS AGENCY (BDA) - PRODUCTS MATRIX PAGE SCRIPT STARTS
//===========================================================================
document.addEventListener("DOMContentLoaded", () => {
    const searchInput = document.getElementById("pharmaLiveSearchInput");
    const tableBody = document.getElementById("pharmaTableBodyRows");
    const noResultsMsg = document.getElementById("noPharmaResultsMessage");
    
    // Actions Targets
    const selectAllCheckbox = document.getElementById("selectAllPharmaRows");
    const exportExcelBtn = document.getElementById("exportExcelBtn");
    const shareWhatsappBtn = document.getElementById("shareWhatsappBtn");
    const checkedRowsCount = document.getElementById("checkedRowsCount");
    
    const categoryItems = document.querySelectorAll(".category-item");
    const companyCheckboxes = document.querySelectorAll(".company-filter-checkbox");

    if (!tableBody) return;

    // 1. रो-डेटा को क्रैश होने से बचाने के लिए सेफ सेलेक्टर फंक्शन
    function getTableRowData(row) {
        if (!row || !row.cells || row.cells.length < 7) {
            return { brand: "", needle: "", suture: "", comp: "", mrp: "" };
        }
        return {
            brand: row.cells[1] ? row.cells[1].innerText.trim() : "",
            needle: row.cells[2] ? row.cells[2].innerText.trim() : "",
            suture: row.cells[3] ? row.cells[3].innerText.trim() : "",
            comp: row.cells[4] ? row.cells[4].innerText.trim() : "",
            mrp: row.cells[5] ? row.cells[5].innerText.trim() : ""
        };
    }

    // 2. 3-लेवल हाइब्रिड मास्टर फ़िल्टर इंजन (खोज + श्रेणी + कंपनी)
    function filterDashboardInventory() {
        const rows = tableBody.getElementsByTagName("tr");
        const searchTerm = searchInput ? searchInput.value.toLowerCase().trim() : "";
        
        const activeCatItem = document.querySelector(".category-item.active-cat");
        const activeCategory = activeCatItem ? activeCatItem.getAttribute("data-category") : "all";

        const selectedCompanies = [];
        companyCheckboxes.forEach(cb => {
            if (cb.checked) selectedCompanies.push(cb.value);
        });

        let matchCount = 0;

        Array.from(rows).forEach(row => {
            const rowCategory = row.getAttribute("data-prod-cat") || "";
            const rowCompany = row.getAttribute("data-company") || "";
            const data = getTableRowData(row);

            const categoryMatches = (activeCategory === "all" || rowCategory === activeCategory);
            const companyMatches = selectedCompanies.includes(rowCompany);
            
            // पूरे रो के किसी भी अक्षर को टाइप करके खोजने की सुविधा
            const rowFullText = `${data.brand} ${data.needle} ${data.suture} ${data.comp}`.toLowerCase();
            const searchMatches = rowFullText.includes(searchTerm);

            if (categoryMatches && companyMatches && searchMatches) {
                row.style.display = "";
                matchCount++;
            } else {
                row.style.display = "none";
                
                // छिपे हुए रो को अनचेक करना ताकि गलत एक्सपोर्ट न हो
                const cb = row.querySelector(".row-select-input");
                if (cb) {
                    cb.checked = false;
                    row.classList.remove("row-is-selected");
                }
            }
        });

        if (noResultsMsg) {
            noResultsMsg.style.display = (matchCount === 0) ? "block" : "none";
        }

        syncActionControls();
    }

    // 3. काउंटर और इनेबल/डिसेबल लॉजिक कंट्रोलर
    function syncActionControls() {
        const allInputs = tableBody.querySelectorAll(".row-select-input");
        const checkedInputs = tableBody.querySelectorAll(".row-select-input:checked");
        const visibleInputs = tableBody.querySelectorAll("tr:not([style*='display: none']) .row-select-input");
        const visibleChecked = tableBody.querySelectorAll("tr:not([style*='display: none']) .row-select-input:checked");

        const count = checkedInputs.length;
        if (checkedRowsCount) checkedRowsCount.innerText = count;

        const isBtnDisabled = (count === 0);
        if (exportExcelBtn) exportExcelBtn.disabled = isBtnDisabled;
        if (shareWhatsappBtn) shareWhatsappBtn.disabled = isBtnDisabled;

        if (selectAllCheckbox) {
            selectAllCheckbox.checked = (visibleInputs.length === visibleChecked.length && visibleInputs.length > 0);
        }
    }

    // 4. 'Select All' चेकबॉक्स पर क्लिक इवेंट
    if (selectAllCheckbox) {
        selectAllCheckbox.addEventListener("change", function () {
            const isChecked = this.checked;
            const rows = tableBody.getElementsByTagName("tr");
            
            Array.from(rows).forEach(row => {
                if (row.style.display !== "none") {
                    const cb = row.querySelector(".row-select-input");
                    if (cb) {
                        cb.checked = isChecked;
                        if (isChecked) row.classList.add("row-is-selected");
                        else row.classList.remove("row-is-selected");
                    }
                }
            });
            syncActionControls();
        });
    }

    // 5. सिंगल चेकबॉक्स क्लिक इवेंट (डेलिगेटेड मोड - सेफ)
    tableBody.addEventListener("change", function (e) {
        if (e.target && e.target.classList.contains("row-select-input")) {
            const tr = e.target.closest("tr");
            if (tr) {
                if (e.target.checked) tr.classList.add("row-is-selected");
                else tr.classList.remove("row-is-selected");
            }
            syncActionControls();
        }
    });

    // फ़िल्टर लिसनर्स बाइंडिंग
    companyCheckboxes.forEach(cb => cb.addEventListener("change", filterDashboardInventory));
    if (searchInput) searchInput.addEventListener("input", filterDashboardInventory);
    
    categoryItems.forEach(item => {
        item.addEventListener("click", function (e) {
            e.preventDefault();
            categoryItems.forEach(cat => cat.classList.remove("active-cat"));
            this.classList.add("active-cat");
            filterDashboardInventory();
        });
    });

    // 6. माइक्रोसॉफ्ट एक्सेल एक्सपोर्ट इंजन (.xlsx)
    if (exportExcelBtn) {
        exportExcelBtn.addEventListener("click", () => {
            const checkedBoxes = tableBody.querySelectorAll(".row-select-input:checked");
            const exportData = [];
            
            // एक्सेल के लिए बिल्कुल सटीक 5 हेडर कॉलम्स
            exportData.push(["BRAND", "NEEDLE'S PROFILE", "SUTURE'S PROFILE", "COMPOSITION", "MRP (BOX/12PCS)"]);

            checkedBoxes.forEach(cb => {
                const tr = cb.closest("tr");
                if (tr) {
                    const data = getTableRowData(tr);
                    exportData.push([data.brand, data.needle, data.suture, data.comp, data.mrp]);
                }
            });

            if (typeof XLSX !== "undefined") {
                const wb = XLSX.utils.book_new();
                const ws = XLSX.utils.aoa_to_sheet(exportData);
                XLSX.utils.book_append_sheet(wb, ws, "Surgical Profiles");
                XLSX.writeFile(wb, "BDA_Surgical_Profiles_Matrix.xlsx");
            } else {
                alert("Excel Engine error: Please check your internet connection to load SheetJS.");
            }
        });
    }

    // 7. एडवांस व्हाट्सएप शेयरिंग मैसेज इंजन
    if (shareWhatsappBtn) {
        shareWhatsappBtn.addEventListener("click", () => {
            const checkedBoxes = tableBody.querySelectorAll(".row-select-input:checked");
            let textMessage = "*BDA GLOBAL MATRIX - SURGICAL PROFILES*\n";
            textMessage += "========================================\n\n";

            let idx = 1;
            checkedBoxes.forEach(cb => {
                const tr = cb.closest("tr");
                if (tr) {
                    const data = getTableRowData(tr);
                    textMessage += `${idx}. *${data.brand}*\n`;
                    textMessage += `   • *Needle:* ${data.needle}\n`;
                    textMessage += `   • *Suture:* ${data.suture}\n`;
                    textMessage += `   • *Composition:* ${data.comp}\n`;
                    textMessage += `   • *MRP:* ${data.mrp}\n`;
                    textMessage += "----------------------------------------\n";
                    idx++;
                }
            });

            textMessage += "\n_Generated via BDA Global System Matrix_";
            window.open(`https://whatsapp.com{encodeURIComponent(textMessage)}`, "_blank");
        });
    }

    // 8. "View More" बटन क्लिक इवेंट
    tableBody.addEventListener("click", function (e) {
        if (e.target && e.target.classList.contains("row-view-details")) {
            e.preventDefault();
            const tr = e.target.closest("tr");
            const data = getTableRowData(tr);
            
            // अगर लॉगिन ब्लॉक पहले से मौजूद है, तो उसे खोलें वरना सेफ अलर्ट दिखाएँ
            alert(`Opening detailed specification metrics for *${data.brand}*.\n\nNeedle Profile: ${data.needle}\nSuture Profile: ${data.suture}`);
        }
    });
});
//=================================================================================
//            BIHAR DRUGS AGENCY (BDA)- PRODUCTS MATRIX PAGE SCRIPT ENDS
//=================================================================================

