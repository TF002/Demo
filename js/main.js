// 票据系统主要JavaScript功能文件
// 包含模板切换、下载、批量生成、随机数字生成、Logo上传等功能

// ==================== Excel Date/Time Formatting Helpers ====================
// 使用 SheetJS SSF 模块来格式化从Excel读取的数字日期/时间
// 这可以正确处理Excel的日期系统（包括1900年闰年问题）
function formatExcelDate(serial) {
    if (typeof serial === 'number' && serial > 0) {
        // XLSX.SSF.format可将序列号转为日期字符串
        return XLSX.SSF.format('mm/dd/yyyy', serial);
    }
    return serial; // 如果不是数字或无效，返回原值
}

function formatExcelTime(serial, withSeconds = true) {
    if (typeof serial === 'number' && serial > 0) {
        // 时间在Excel中是0到1之间的小数，但日期时间混合的数字也可能存在
        const format = withSeconds ? 'hh:mm:ss' : 'hh:mm';
        return XLSX.SSF.format(format, serial);
    }
    return serial;
}

// ==================== 模板切换逻辑 ====================
document.addEventListener('DOMContentLoaded', function () {
    const templateSelector = document.getElementById('template-selector');
    if (templateSelector) {
        templateSelector.addEventListener('change', function () {
            const selectedValue = this.value;
            const allTemplates = document.querySelectorAll('.receipt-template');

            allTemplates.forEach(template => {
                if (template.id === selectedValue) {
                    template.style.display = '';
                    // 当切换模板时，如果已有上传的logo，则应用它
                    if (currentUploadedLogo) {
                        const logoImg = template.querySelector('.logo img');
                        if (logoImg) {
                            logoImg.src = currentUploadedLogo;
                        }
                    }
                    // 为当前显示的模板生成条形码
                    const barcodeElement = template.querySelector('.barcode');
                    console.log(`切换到模板 ${template.id}，条形码元素:`, barcodeElement);
                    if (barcodeElement) {
                        ensureJsBarcode(() => {
                            applyRandomBarcode(barcodeElement);
                        });
                    }
                } else {
                    template.style.display = 'none';
                }
            });

            // 如果切换到Grocery模板，重新生成随机数字和更新商品数量
            if (selectedValue === 'grocery-furniture-store-for-online-receipts') {
                setTimeout(() => {
                    generateGroceryRandomNumbers();
                    updateGroceryItemCount();
                }, 100);
            }

            // 如果切换到Electronic Store模板，重新生成随机数字
            if (selectedValue === 'electronic-store-receipt-maker') {
                setTimeout(() => {
                    generateElectronicStoreRandomNumbers();
                    // 生成条形码
                    const barcodeElement = document.querySelector('#electronic-store-receipt-maker .barcode');
                    if (barcodeElement) {
                        ensureJsBarcode(() => {
                            applyRandomBarcode(barcodeElement);
                        });
                    }
                }, 100);
            }

            // 如果切换到Online Furniture Shop模板，重新生成随机数字
            if (selectedValue === 'online-receipts-for-furniture-shop') {
                setTimeout(() => {
                    generateOnlineFurnitureShopRandomNumbers();
                    // 生成条形码
                    const barcodeElement = document.querySelector('#online-receipts-for-furniture-shop .barcode');
                    if (barcodeElement) {
                        ensureJsBarcode(() => {
                            applyRandomBarcode(barcodeElement);
                        });
                    }
                    // 更新商品总数
                    updateOnlineFurnitureShopItemCount();
                }, 100);
            }
        });
    }
});

// ==================== 单张票据下载功能 ====================
document.addEventListener('DOMContentLoaded', function () {
    const downloadBtn = document.getElementById('download-btn');
    if (downloadBtn) {
        downloadBtn.addEventListener('click', async function () {
            try {
                const canvas = await generateReceiptCanvasSafe();
                const blob = await new Promise((resolve, reject) =>
                    canvas.toBlob((b) => b ? resolve(b) : reject(new Error('toBlob 失败')), 'image/png')
                );
                const link = document.createElement('a');
                const visibleTemplate = document.querySelector('.receipt-template:not([style*="display:none"]):not([style*="display: none"])') ||
                    document.getElementById('retail-electronic-home-Improvements-chain-store');
                link.href = URL.createObjectURL(blob);
                link.download = `${visibleTemplate ? visibleTemplate.id : 'receipt'}.png`;
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                setTimeout(() => URL.revokeObjectURL(link.href), 0);
            } catch (error) {
                console.error('生成或下载票据图片时出错:', error);
                alert('图片导出失败：请确保图片/Logo来源允许导出，或改用上传Logo/通过本地HTTP服务访问。');
            }
        });
    }
});


// ==================== 批量生成功能 ====================

// 新增：安全截图函数，专门用于单张下载，在file://环境下强制无图导出
async function generateReceiptCanvasSafe() {
    // 查找当前可见的模板进行截图
    let receiptElement = document.querySelector('.receipt-template:not([style*="display: none"])') || document.getElementById('retail-electronic-home-Improvements-chain-store');

    // 如果找不到可见的模板，尝试根据模板选择器来确定
    if (!receiptElement || receiptElement.style.display === 'none') {
        const templateSelector = document.getElementById('template-selector');
        if (templateSelector) {
            const selectedValue = templateSelector.value;
            const targetTemplate = document.getElementById(selectedValue);
            if (targetTemplate) {
                console.log('根据选择器找到模板:', selectedValue);
                receiptElement = targetTemplate;
            }
        }
    }

    // 统一目标为内部的 .card.xrt 元素，与单张下载保持一致
    const targetElementForCanvas = receiptElement.querySelector('.card.xrt');

    if (targetElementForCanvas) {
        return await generateCanvasForTemplateSafe(targetElementForCanvas);
    }

    // 如果模板内部没有 .card.xrt，则回退到截取整个模板
    console.warn("在模板中未找到 .card.xrt 元素，将截取整个模板。");
    return await generateCanvasForTemplateSafe(receiptElement);
}

// 安全版本的Canvas生成函数，专门处理file://环境
async function generateCanvasForTemplateSafe(receiptElement) {
    console.log('安全模式截图 - 目标元素:', receiptElement);
    
    // 确保元素可见
    if (receiptElement.style.display === 'none') {
        receiptElement.style.display = 'block';
    }

    // 等待元素渲染
    await new Promise(resolve => setTimeout(resolve, 100));

    const originalWidth = receiptElement.offsetWidth;
    const targetWidth = 720;
    const scale = originalWidth > 0 ? targetWidth / originalWidth : 1;

    console.log('原始宽度:', originalWidth, '目标宽度:', targetWidth, '缩放比例:', scale);

    // 关键：在file://环境下，临时隐藏所有可能污染画布的图片
    const isFileProtocol = window.location && window.location.protocol === 'file:';
    const temporarilyHiddenElements = [];

    if (isFileProtocol) {
        console.log('检测到file://环境，启用安全模式：保留data URL图片，隐藏其他图片');
        
        // 隐藏所有img元素，但保留data URL和blob URL格式的图片（如上传的Logo）
        const allImages = receiptElement.querySelectorAll('img');
        allImages.forEach(img => {
            const isDataUrl = img.src && (img.src.startsWith('data:') || img.src.startsWith('blob:'));
            if (!isDataUrl && img.style.display !== 'none') {
                temporarilyHiddenElements.push({ el: img, display: img.style.display });
                img.style.display = 'none';
            }
        });

        // 移除所有背景图片
        const allElements = receiptElement.querySelectorAll('*');
        allElements.forEach(el => {
            const computedStyle = window.getComputedStyle(el);
            if (computedStyle.backgroundImage && computedStyle.backgroundImage !== 'none') {
                temporarilyHiddenElements.push({ el: el, backgroundImage: el.style.backgroundImage });
                el.style.backgroundImage = 'none';
            }
        });
    }

    const parentElement = receiptElement.parentElement;
    const originalParentHeight = parentElement.style.height;

    try {
        parentElement.style.height = 'auto';

        // 等待字体加载
        await document.fonts.ready;
        console.log("All fonts are loaded and ready for html2canvas.");

        const options = {
            scale: scale,
            useCORS: false, // 在安全模式下禁用CORS
            allowTaint: false,
            backgroundColor: '#ffffff',
            logging: true,
            onclone: (clonedDoc) => {
                // 在克隆文档中也确保无图片，但保留data URL格式的图片
                if (isFileProtocol) {
                    clonedDoc.querySelectorAll('img').forEach(img => { 
                        const isDataUrl = img.src && (img.src.startsWith('data:') || img.src.startsWith('blob:'));
                        if (!isDataUrl) {
                            img.style.display = 'none'; 
                        }
                    });
                    clonedDoc.querySelectorAll('*').forEach(el => {
                        el.style.backgroundImage = 'none';
                    });
                }

                // 生成条形码
                const barcodeElement = clonedDoc.querySelector('.barcode');
                if (barcodeElement && typeof JsBarcode !== 'undefined') {
                    applyRandomBarcode(barcodeElement);
                }
            },
        };

        const canvas = await html2canvas(receiptElement, options);
        console.log('安全模式Canvas生成成功，尺寸:', canvas.width, 'x', canvas.height);

        if (canvas.width === 0 || canvas.height === 0) {
            throw new Error('生成的Canvas尺寸为0，可能是元素不可见或样式问题');
        }

        return canvas;
    } catch (error) {
        console.error('安全模式html2canvas生成失败:', error);
        throw error;
    } finally {
        // 恢复所有被隐藏的元素
        parentElement.style.height = originalParentHeight;
        temporarilyHiddenElements.forEach(({ el, display, backgroundImage }) => {
            if (display !== undefined) el.style.display = display;
            if (backgroundImage !== undefined) el.style.backgroundImage = backgroundImage;
        });
    }
}

// 新增：用于生成单个票据Canvas的辅助函数
async function generateReceiptCanvas() {
    // 查找当前可见的模板进行截图
    let receiptElement = document.querySelector('.receipt-template:not([style*="display: none"])') || document.getElementById('retail-electronic-home-Improvements-chain-store');

    // 如果找不到可见的模板，尝试根据模板选择器来确定
    if (!receiptElement || receiptElement.style.display === 'none') {
        const templateSelector = document.getElementById('template-selector');
        if (templateSelector) {
            const selectedValue = templateSelector.value;
            const targetTemplate = document.getElementById(selectedValue);
            if (targetTemplate) {
                console.log('根据选择器找到模板:', selectedValue);
                receiptElement = targetTemplate;
            }
        }
    }

    // 统一目标为内部的 .card.xrt 元素，与单张下载保持一致
    const targetElementForCanvas = receiptElement.querySelector('.card.xrt');

    if (targetElementForCanvas) {
        return await generateCanvasForTemplate(targetElementForCanvas);
    }

    // 如果模板内部没有 .card.xrt，则回退到截取整个模板
    console.warn("在模板中未找到 .card.xrt 元素，将截取整个模板。");
    return await generateCanvasForTemplate(receiptElement);
}

async function generateCanvasForTemplate(receiptElement) {
    console.log('目标元素:', receiptElement);
    console.log('元素可见性:', receiptElement.style.display);
    console.log('元素尺寸:', receiptElement.offsetWidth, 'x', receiptElement.offsetHeight);

    // 确保元素可见
    if (receiptElement.style.display === 'none') {
        receiptElement.style.display = 'block';
    }

    // console.log(2222)
    // return;

    // 等待元素渲染
    await new Promise(resolve => setTimeout(resolve, 100));

    const originalWidth = receiptElement.offsetWidth;
    const targetWidth = 720;
    const scale = originalWidth > 0 ? targetWidth / originalWidth : 1;

    console.log('原始宽度:', originalWidth, '目标宽度:', targetWidth, '缩放比例:', scale);
    // 等待所有图片加载完成
    const images = receiptElement.querySelectorAll('img');
    const imagePromises = Array.from(images).map(img => {
        return new Promise((resolve) => {
            if (img.complete && img.naturalHeight !== 0) {
                resolve();
            } else {
                img.onload = resolve;
                img.onerror = resolve; // 即使失败也继续
            }
        });
    });
    await Promise.all(imagePromises);

    // 关键改动：通过临时修改父元素样式，让内容自然撑开，再由html2canvas自动计算高度
    const parentElement = receiptElement.parentElement;
    const originalParentHeight = parentElement.style.height;

    // file:// 场景下为避免跨域污染（tainted canvas），对非 data: 的图片进行临时隐藏
    const isFileProtocol = window.location && window.location.protocol === 'file:';
    const temporarilyHiddenImages = [];

    try {
        parentElement.style.height = 'auto'; // 允许子元素撑开父元素

        // 关键改动：在截图前，确保所有字体已加载完成
        await document.fonts.ready;
        console.log("All fonts are loaded and ready for html2canvas.");

        const options = {
            scale: scale,
            useCORS: true,
            // 关键：禁止允许污染，这样一旦有跨域资源会阻止绘制，避免最终 toBlob 报错
            allowTaint: false,
            backgroundColor: '#ffffff',
            logging: true,
            // 移除固定的 width 和 height，让 html2canvas 自动测量
            onclone: (clonedDoc) => {
                // 在 file:// 环境下，彻底移除克隆文档中的潜在污染源
                if (isFileProtocol) {
                    // 隐藏所有 img
                    clonedDoc.querySelectorAll('img').forEach(img => { img.style.display = 'none'; });
                    // 移除所有背景图
                    clonedDoc.querySelectorAll('*').forEach(el => {
                        const cs = clonedDoc.defaultView.getComputedStyle(el);
                        if (cs && cs.backgroundImage && cs.backgroundImage !== 'none') {
                            el.style.backgroundImage = 'none';
                        }
                    });
                }
                const barcodeElement = clonedDoc.querySelector('.barcode');
                if (barcodeElement && typeof JsBarcode !== 'undefined') {
                    applyRandomBarcode(barcodeElement);
                }

                // 确保克隆文档中的图片也能正确显示
                const clonedImages = clonedDoc.querySelectorAll('img');
                clonedImages.forEach(img => {
                    if (img.src && img.src.startsWith('data:')) {
                        img.style.display = 'block';
                    }
                });

                // 确保Logo居中样式正确应用
                const logoElements = clonedDoc.querySelectorAll('.logo');
                logoElements.forEach(logo => {
                    logo.style.textAlign = 'center';
                    const logoImg = logo.querySelector('img');
                    if (logoImg) {
                        logoImg.style.display = 'block';
                        logoImg.style.margin = '0 auto';
                        logoImg.style.maxWidth = '280px';
                        logoImg.style.maxHeight = '80px';
                        logoImg.style.objectFit = 'contain';
                    }
                });

                // 确保表格居中样式正确应用
                const tables = clonedDoc.querySelectorAll('table');
                tables.forEach(table => {
                    if (table.classList.contains('table-borderless')) {
                        table.style.margin = '0 auto';
                        table.style.width = '100%';
                    }
                });
            },
        };

        // 在 file:// 环境下，临时隐藏所有非 data: 与非 blob: 的图片，避免画布被污染
        if (isFileProtocol) {
            Array.from(images).forEach((imgEl) => {
                const src = imgEl.getAttribute('src') || '';
                if (src && !src.startsWith('data:') && !src.startsWith('blob:')) {
                    temporarilyHiddenImages.push({ el: imgEl, display: imgEl.style.display });
                    imgEl.style.display = 'none';
                }
            });
        }

        const canvas = await html2canvas(receiptElement, options);
        console.log('Canvas生成成功，尺寸:', canvas.width, 'x', canvas.height);

        if (canvas.width === 0 || canvas.height === 0) {
            throw new Error('生成的Canvas尺寸为0，可能是元素不可见或样式问题');
        }

        return canvas;
    } catch (error) {
        console.error('html2canvas生成失败:', error);
        throw error;
    } finally {
        // 无论成功或失败，都恢复父元素的原始样式
        parentElement.style.height = originalParentHeight;
        // 恢复被临时隐藏的图片
        if (temporarilyHiddenImages.length) {
            temporarilyHiddenImages.forEach(({ el, display }) => { el.style.display = display; });
        }
    }
}

document.addEventListener('DOMContentLoaded', function () {
    const batchGenerateBtn = document.getElementById('batch-generate-btn');
    if (batchGenerateBtn) {
        batchGenerateBtn.addEventListener('click', async function () {
            const fileInput = document.getElementById('excel-file-input');
            const statusDisplay = document.getElementById('status-display');

            // 支持三个模板的批量生成
            const selectedTemplate = document.getElementById('template-selector').value;

            if (fileInput.files.length === 0) {
                alert('请先选择一个 Excel 文件！');
                return;
            }

            const file = fileInput.files[0];
            // 获取Excel文件名（去除扩展名）
            const excelFileName = file.name.replace(/\.[^/.]+$/, "");
            const reader = new FileReader();

            reader.onload = async function (e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const sheet = workbook.Sheets[firstSheetName];
                    // 关键改动：将 raw 设置为 true，以读取原始单元格数据，避免大数字格式化问题
                    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });

                    // 调试：检查原始Excel数据
                    console.log('原始Excel行数据（前3行）:', rows.slice(0, 3));
                    console.log('第一行数据长度:', rows[0] ? rows[0].length : 0);
                    console.log('第二行数据长度:', rows[1] ? rows[1].length : 0);
                    if (rows[1] && rows[1].length > 13) {
                        console.log('第二行N列（索引13）原始数据:', rows[1][13]);
                    }
                    if (rows.length < 2) {
                        statusDisplay.textContent = '错误：Excel文件至少需要包含一个标题行和一行数据。';
                        return;
                    }

                    const headers = rows.shift(); // 第一行作为标题
                    const receiptsData = [];
                    rows.forEach(row => {
                        const rowData = {};

                        headers.forEach((header, index) => {
                            const colName = XLSX.utils.encode_col(index); // A, B, C...
                            const cellValue = row[index];
                            // 读取原始数据，不立即转换为字符串，以便后续对时间和日期等数字进行格式化
                            // 只使用列的字母代号作为key，避免重复的列标题导致数据覆盖问题
                            rowData[colName] = cellValue;
                        });

                        // 手动添加N列数据（索引13）
                        if (row.length > 13) {
                            rowData['N'] = row[13];
                        }
                        if(rowData['A']){
                            receiptsData.push(rowData);
                        }
                    });

                    const zip = new JSZip();
                    let generatedCount = 0;
                    console.log(`总共读取到 ${receiptsData.length} 行数据，开始生成票据...`,receiptsData);

                    for (let i = 0; i < receiptsData.length; i++) {
                        const rowData = receiptsData[i];
                        statusDisplay.textContent = `正在生成第 ${i + 1} / ${receiptsData.length} 张...`;

                        await populateReceiptByTemplate(rowData, selectedTemplate);

                        // 使用 requestAnimationFrame 确保在下一次绘制时才截图
                        // 修复：如果 generateReceiptCanvas 抛错，必须 reject，否则会导致永远 pending 而"卡住"
                        // 批量生成也使用安全模式，避免 tainted canvas 错误
                        const canvas = await new Promise((resolve, reject) => {
                            requestAnimationFrame(() => {
                                generateReceiptCanvasSafe()
                                    .then(resolve)
                                    .catch((err) => {
                                        console.error('生成Canvas失败:', err);
                                        reject(err);
                                    });
                            });
                        });

                        const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/png'));

                        const fileName = `${excelFileName}-${i + 1}.png`;
                        zip.file(fileName, blob);
                        generatedCount++;
                    }

                    statusDisplay.textContent = '正在打包ZIP文件...';
                    const zipBlob = await zip.generateAsync({ type: "blob" });

                    const link = document.createElement('a');
                    link.href = URL.createObjectURL(zipBlob);
                    link.download = `${excelFileName}-receipts.zip`;
                    document.body.appendChild(link);
                    link.click();
                    document.body.removeChild(link);


                    statusDisplay.textContent = `所有票据已生成并打包下载！共生成 ${generatedCount} 张票据。`;
                } catch (error) {
                    console.error('处理Excel文件时出错:', error);
                    statusDisplay.textContent = '处理文件失败，请检查文件格式或控制台信息。';
                    alert('处理文件失败，请确保文件格式正确！');
                }
            };

            reader.readAsArrayBuffer(file);
        });
    }
});

// 根据模板类型调用相应的数据填充函数
async function populateReceiptByTemplate(data, templateType) {
    console.log('当前模版--',templateType)
    // 根据模板类型选择不同的填充函数
    if (templateType === 'retail-electronic-home-Improvements-chain-store') {
        await populateRetailReceipt(data);
    } else if (templateType === 'grocery-furniture-store-for-online-receipts') {
        await populateGroceryReceipt(data);
    } else if (templateType === 'electronic-store-receipt-maker') {
        await populateElectronicStoreReceipt(data);
    } else if (templateType === 'general-grocery-store-template-for-food-grocery-meat-juices-and-bread-receipts') {
        await populateGeneralGroceryReceipt(data);
    } else if (templateType === 'online-receipts-for-furniture-shop') {
        await populateOnlineFurnitureShopReceipt(data);
    }
    // ... 可以为其他模板添加更多 else if
}

async function populateGeneralGroceryReceipt(data) {
    // 获取当前模板的 DOM 元素
    const template = document.getElementById('general-grocery-store-template-for-food-grocery-meat-juices-and-bread-receipts');
    if (!template) return;
    console.log('template',template)
    // --- 清理旧数据 ---
    // 隐藏所有商品行并清除内容
    const itemRows = template.querySelectorAll('.added_content.Pitem .remove_tr');
    itemRows.forEach(row => {
        row.style.display = 'none';
        row.querySelectorAll('span[class*="itemname"], span[class*="itemprice"], span[class*="itemF"], span[class*="itemdown"]').forEach(span => {
            span.textContent = '';
        });
    });

    // --- 填充静态数据 ---
    template.querySelector('.enddate').textContent = formatExcelDate(data['A']) || '09/11/2025';
    template.querySelector('.endtime').textContent = formatExcelTime(data['B']) || '15:03';

    template.querySelector('.address1').textContent = data['E'] || '804-330-7365';
    template.querySelector('.address2').textContent = data['F'] || '7107 Forest Hill Ave';
    template.querySelector('.address3').textContent = data['G'] || 'RichmondVA 23225';

    // --- 动态处理商品和计算 ---
    let subtotal = 0;
    for (let i = 1; i <= 10; i++) { // 循环10个商品
        const nameCol = XLSX.utils.encode_col(7 + (i - 1) * 5);  // H, L, P...
        const descCol = XLSX.utils.encode_col(8 + (i - 1) * 5);  // I, M, Q...
        const priceCol = XLSX.utils.encode_col(9 + (i - 1) * 5); // J, N, R...
        const upcCol = XLSX.utils.encode_col(10 + (i - 1) * 5);  // K, O, S...
        const tagCol = XLSX.utils.encode_col(11 + (i - 1) * 5);  // K, O, S...

        // --- 新增调试代码 ---
        if (i === 5) {
            console.log("--- 调试模板4，第5个商品 ---");
            console.log(`品名列应为: X (计算结果: ${nameCol})`);
            console.log(`从Excel X列读取到的原始值是:`, data[nameCol]);
            console.log(`该值的数据类型是: ${typeof data[nameCol]}`);
        }
        // --- 调试代码结束 ---

        const itemName = data[nameCol];
        const itemPrice = parseFloat(data[priceCol]);

        if (itemName && !isNaN(itemPrice)) {
            const row = template.querySelector(`.Pitem .remove_tr:nth-child(${i})`);
            if (row) {
                row.style.display = 'table-row';
                row.querySelector(`.itemname${i}`).textContent = itemName;
                row.querySelector(`.itemdown${i}`).textContent = data[descCol] || '';
                row.querySelector(`.itemprice${i}`).textContent = itemPrice.toFixed(2);
                row.querySelector(`.itemF${i}`).textContent = data[tagCol] || 'F';
                subtotal += itemPrice;
            }
        }
    }

    // --- 计算和填充总计 ---
    const taxValue = parseFloat(data['AW']) || 0; // 税费金额从 AW 列读取
    const total = subtotal + taxValue;

    template.querySelector('.subtotal').textContent = subtotal.toFixed(2);
    // template.querySelector('.taxname1').textContent = data['AV'] || 'Tax'; // 税费标签从 AV 列读取
    template.querySelector('.taxprice1').textContent = taxValue.toFixed(2);
    template.querySelector('.total').textContent = total.toFixed(2);
    template.querySelector('.totalceil').textContent = total.toFixed(2);
    template.querySelector('.remain').textContent = '0.00';


    // --- 填充支付和消息 ---
    // 安全地填充每个元素，检查是否存在
    const cardTypeMain = data['AX'] || 'VISA';
    const cardTypeSecondary = data['AY'] || cardTypeMain;

    const swipeCashEl = template.querySelector('.swipe.cash');
    if (swipeCashEl) swipeCashEl.textContent = cardTypeMain;

    const cardTypeElements = template.querySelectorAll('.cardtype');
    cardTypeElements.forEach(el => {
        if (el) el.textContent = cardTypeSecondary;
    });

    const acctEl = template.querySelector('.acct');
    if (acctEl) acctEl.textContent = data['AZ'] || '6587';

    // BA列: 映射到 message1
    const message1El = template.querySelector('.message1');
    if (message1El) message1El.textContent = data['BJ'] || "Together, we'll get through this COVID-19 Situation.";

    // BB列: 映射到 message2
    const message2El = template.querySelector('.message2');
    if (message2El) message2El.textContent = '★★★ CUSTOMER COPY ★★★';

    // --- 处理Logo (假设为BC列) ---
    const logoImg = template.querySelector('.logo img');
    console.log('data123,BC', data);
    const logoFilename = data['BK'];
    if (logoFilename) {
        logoImg.src = `./logo/${logoFilename}`;
        logoImg.style.display = 'inline';
    } else {
        // 如果没有提供logo，可以隐藏img标签或显示占位符
        logoImg.style.display = 'none';
    }
}

// Retail模板数据填充函数（原有逻辑）
async function populateRetailReceipt(data) {
    const templateId = 'retail-electronic-home-Improvements-chain-store';
    const template = document.getElementById(templateId);
    if (!template) return;
    console.log('template1',template)
    // --- 清理旧数据 ---
    const itemRows = template.querySelectorAll('.Pitem .remove_tr');
    itemRows.forEach(row => {
        row.style.display = 'none';
        const cells = row.querySelectorAll('.subItemTr td');
        if (cells.length === 4) {
            cells[0].querySelector('span').textContent = '';
            cells[1].querySelector('span').textContent = '';
            cells[2].querySelector('span').textContent = '';
            cells[3].querySelectorAll('span')[1].textContent = '';
        }
    });

    console.log('itemRows',itemRows,'data',data)

    // // --- 填充静态数据 ---
    // 时间
    template.querySelector('.checkindate').textContent = formatExcelDate(data['A']) || '';
    template.querySelector('.checkintime').textContent = formatExcelTime(data['B']) || '';
    
    template.querySelector('.address1').textContent = data['F'] || '';
    template.querySelector('.address2').textContent = data['G'] || '';
    template.querySelector('.address3').textContent = data['BA'] || '';
    template.querySelector('.cityName').textContent = data['E'] || '';
    // // I 列不再映射到 .recVcd

    // --- 动态处理商品和计算 ---
    let subtotal = 0;
    // 商品从第 I 列 (索引 8) 开始，每组 4 列
    for (let i = 1; i <= 10; i++) {
        // I=8, J=9, K=10, L=11 priceCol
        const nameCol = XLSX.utils.encode_col(7 + (i - 1) * 4);
        const codeCol = XLSX.utils.encode_col(8 + (i - 1) * 4);
        const priceCol = XLSX.utils.encode_col(9 + (i - 1) * 4);
        const typeCol = XLSX.utils.encode_col(10 + (i - 1) * 4);

        const itemName = data[nameCol];
        const itemPrice = parseFloat(data[priceCol]);
        console.log(3333,{
            codeCol,
            nameCol,
            typeCol,
            priceCol,
            itemPrice,
            itemName
        })

        if (itemName && !isNaN(itemPrice)) {
            const row = itemRows[i - 1];
            console.log('row22',row)
            if (row) {
                row.style.display = 'table-row';
                const cells = row.querySelectorAll('.subItemTr td');
                console.log('cells',cells)
                cells[0].querySelector('span').textContent = data[codeCol] || '';
                cells[1].querySelector('span').textContent = itemName;
                cells[2].querySelector('span').textContent = data[typeCol] || '';
                cells[3].querySelectorAll('span')[1].textContent = itemPrice.toFixed(2);
                subtotal += itemPrice;
            }
        }
    }

    // // --- 计算和填充总计 ---
    const taxValue = parseFloat(data['AV']) || 0;
    const total = subtotal + taxValue;

    // --- 使用更安全的方式更新总计 ---
    // Subtotal - 注意：小计是在itemTable的第二个tbody中
    const subtotalElement = template.querySelector('.itemTable > tbody:not(.Pitem) td[style*="width:30%"] span:last-child');
    if (subtotalElement) {
        subtotalElement.textContent = subtotal.toFixed(2);
    }

    // Tax
    // const taxLabelElement = template.querySelector('#taxLabel');
    // if (taxLabelElement) {
    //     taxLabelElement.textContent = data['AY'] || 'TAX';
    // }
    const taxValueElement = template.querySelector('#taxValue');
    if (taxValueElement) {
        taxValueElement.textContent = taxValue.toFixed(2);
    }

    const totalval2 = template.querySelector('#totalval2');
    if (totalval2) {
        totalval2.textContent = total.toFixed(2);
    }

     const totalval1 = template.querySelector('#totalval1');
    if (totalval1) {
        totalval1.textContent = total.toFixed(2);
    }
    function generateRandomNumber2(length) {
        if (length <= 0) return '';
        let firstDigit = Math.floor(Math.random() * 9) + 1;
        if (length === 1) return firstDigit.toString();
        let restOfDigits = Array.from({ length: length - 1 }, () => Math.floor(Math.random() * 10)).join('');
        return firstDigit + restOfDigits;
    }

    const recVcdElement = document.querySelector('.recVcd');
    if (recVcdElement) {
        const recPart = `REC#${generateRandomNumber2(1)}-${generateRandomNumber2(4)}-${generateRandomNumber2(4)}-${generateRandomNumber2(4)}-${generateRandomNumber2(4)}-${generateRandomNumber2(1)}`;
        const vcdPart = `VCD#${generateRandomNumber2(3)}-${generateRandomNumber2(3)}-${generateRandomNumber2(3)}`;
        recVcdElement.textContent = `${recPart} ${vcdPart}`;
    }


    // // Total & Charge Total - 第三个table用于总计
    // const totalTable = template.querySelectorAll('table')[4]; // 第五个table
    // if (totalTable) {
    //     // const totalElements = totalTable.querySelectorAll('td[style*="width:30%"] span:last-child');
    //     // console.log('totalElements',totalElements)
    //     // if (totalElements.length >= 2) {
    //     //     totalElements[0].textContent = total.toFixed(2); // TOTAL
    //     //     totalElements[1].textContent = total.toFixed(2); // VISA CHARGE
    //     // }
    // }

    // --- 填充支付和消息 ---
    template.querySelector('#cardType').textContent = data['AW'] || '';
    template.querySelector('#cardLastFour').textContent = data['AX'] || '';
    template.querySelector('.message').textContent = data['AY'] || '';

    // --- 处理Logo ---
    const logoImg = template.querySelector('.logo img');
    const logoFilename = data['AZ'];
    console.log('data123,BB', data);

    if (logoFilename) {
        logoImg.src = `./logo/${logoFilename}`;
        logoImg.style.display = 'inline';
        logoImg.parentElement.style.display = 'table-cell';
    } else {
        logoImg.style.display = 'none';
        logoImg.parentElement.style.display = 'none';
    }
}

// Grocery模板数据填充函数
async function populateGroceryReceipt(data) {
    const templateId = 'grocery-furniture-store-for-online-receipts';
    const template = document.getElementById(templateId);
    if (!template) return;

    // --- 清理旧数据 ---
    const itemRows = template.querySelectorAll('.Pfare .remove_tr');
    itemRows.forEach((row, index) => {
        row.style.display = 'none';
        const i = index + 1;
        row.querySelector(`.farename${i}`).textContent = '';
        row.querySelector(`.farenumber${i}`).textContent = '';
        row.querySelector(`.fareprice${i}`).textContent = '';
        row.querySelector(`.fareF${i}`).textContent = '';
        row.querySelector(`.faredown${i}`).textContent = '';
    });

    // --- 填充静态数据 ---
    // template.querySelector('.enddate').textContent = formatExcelDate(data['A']) || '';
    // template.querySelector('.endtime').textContent = formatExcelTime(data['B'], false) || '';
    template.querySelector('.address1').textContent = data['F'] || '';
    template.querySelector('.address2').textContent = data['G'] || '';
    template.querySelector('.address3').textContent = data['H'] || '';
    template.querySelector('#creadId1').textContent = data['E'] || '';

     // --- 处理Logo（调整后的列映射） ---
    // BF列：logo
    const logoImg = template.querySelector('.logo img');
    const logoFilename = data['BV'];
    console.log('data123,BF', data);

    if (logoFilename) {
        logoImg.src = `./logo/${logoFilename}`;
        logoImg.style.display = 'inline';
    } else {
        logoImg.style.display = 'none';
    }

    // --- 动态处理商品和计算 ---
    let subtotal = 0;
    let lastItemIndex = -1;
    for (let i = 1; i <= 10; i++) {
        // I=8, J=9, K=10, L=11
        const nameCol = XLSX.utils.encode_col(8 + (i - 1) * 6);
        // const numberCol = XLSX.utils.encode_col(9 + (i - 1) * 6);
        const priceCol = XLSX.utils.encode_col(10 + (i - 1) * 6); // K列 - 价格
        const tagCol = XLSX.utils.encode_col(11 + (i - 1) * 6);   // L列 - 标签'X'
        const upcCode = XLSX.utils.encode_col(12 + (i - 1) * 6);   // M列 - upc码
        const priceTag = XLSX.utils.encode_col(13 + (i - 1) * 6);   // N列 - 钱标

        const itemName = data[nameCol];
        const itemPrice = parseFloat(data[priceCol]);

        console.log(3333333,{
            nameCol,
            // numberCol,
            priceCol,
            tagCol,
            itemName,
            itemPrice,
            upcCode,
            priceTag
        })

        if (itemName && !isNaN(itemPrice)) {
            const row = itemRows[i - 1];
            if (row) {
                row.style.display = 'table-row';
                row.querySelector(`.farename${i}`).textContent = itemName;
                row.querySelector(`.fareF${i}`).textContent = data[tagCol] || '';
                row.querySelector(`.farenumber${i}`).textContent = data[upcCode] || '';
                row.querySelector(`.fareprice${i}`).textContent = itemPrice.toFixed(2);
                row.querySelector(`.fareX${i}`).textContent = data[priceTag] || '';

                // // 根据用户的描述，L列是'X'标签，对应 .fareX{i}
                // const tagElement = row.querySelector(`.fareX${i}`);
                // if (tagElement) {
                //     tagElement.textContent = data[tagCol] || 'X';
                // }

                subtotal += itemPrice;
                lastItemIndex = i;
            }
        }
    }
    template.querySelector('.subtotal').textContent = subtotal.toFixed(2);
    template.querySelector('.taxpro1').textContent = (data['BQ'] * 100).toFixed(2) || '0.00';
    template.querySelector('.taxprice1').textContent = (data['BR']).toFixed(2) || '0.00';


    // // 处理商品描述
    // const itemDescription = data['AW'];
    // if (itemDescription && lastItemIndex !== -1) {
    //     const descSpan = template.querySelector(`.faredown${lastItemIndex}`);
    //     if (descSpan) {
    //         descSpan.textContent = itemDescription;
    //         descSpan.closest('.down_content').style.display = 'table-row';
    //     }
    // }

    // // --- 计算和填充总计 ---
    const taxValue = parseFloat(data['BR']) || 0;
    const total = subtotal + taxValue;

    // template.querySelector('.subtotal').textContent = subtotal.toFixed(2);
    // template.querySelector('.taxname1').textContent = data['AX'] || 'TAX';
    // template.querySelector('.taxpro1').textContent = data['AY'] || '0.00';
    // template.querySelector('.taxprice1').textContent = taxValue.toFixed(2);
    template.querySelector('.total').textContent = total.toFixed(2);
    template.querySelector('.tend').textContent = total.toFixed(2);

    // // --- 填充支付和消息 ---
    template.querySelector('.swipe').textContent = data['BS'] || '';
    template.querySelector('.lastnumber').textContent = data['BT'] || '';
    template.querySelector('.approval').textContent =  Math.floor(100000 + Math.random() * 900000);
    template.querySelector('.ref').textContent =  Math.floor(100000000000 + Math.random() * 900000);
    template.querySelector('.terminal').textContent =  Math.floor(1000000000 + Math.random() * 900000);

    // // --- 填充消息（调整后的列映射） ---
    // // BD列：Thank You for Shopping With Us!
    // const commentElement = template.querySelector('.comment');
    // if (commentElement) {
    //     commentElement.textContent = data['BD'] || 'Thank You for Shopping With Us!';
    // }

    // // BE列：✯✯✯ CUSTOMER COPY ✯✯✯
    // const messageElement = template.querySelector('.message');
    // if (messageElement) {
    //     messageElement.textContent = data['AZ'] || '✯✯✯ CUSTOMER COPY ✯✯✯';
    // }

    // // --- 更新售卖商品数量 ---
    const itemsSoldElement = template.querySelector('.itemnum');
    if (itemsSoldElement) {
        itemsSoldElement.textContent = lastItemIndex > 0 ? lastItemIndex : 0;
    }

    function generateNumberGroups() {
        let groups = [];
        for (let i = 0; i < 5; i++) {
            // 生成 0~9999 的随机数，补足 4 位
            let group = Math.floor(Math.random() * 10000).toString().padStart(4, '0');
            groups.push(group);
        }
        return groups.join(' ');
    }

    template.querySelector('.barvalue').textContent =  generateNumberGroups();


    // // --- 处理Logo（调整后的列映射） ---
    // // BF列：logo
    // const logoImg = template.querySelector('.logo img');
    // const logoFilename = data['BF'];
    // console.log('data123,BF', data);

    // if (logoFilename) {
    //     logoImg.src = `./logo/${logoFilename}`;
    //     logoImg.style.display = 'inline';
    // } else {
    //     logoImg.style.display = 'none';
    // }
}

// Electronic Store Receipt Maker模板数据填充函数
async function populateElectronicStoreReceipt(data) {
    const templateId = 'electronic-store-receipt-maker';
    const template = document.getElementById(templateId);
    if (!template) return;

    // --- 清理旧数据 ---
    const itemRows = template.querySelectorAll('.Pitem .remove_tr');
    itemRows.forEach((row, index) => {
        row.style.display = 'none';
        const i = index + 1;
        row.querySelector(`.itemCouponA${i}`).textContent = '';
        row.querySelector(`.itemCouponB${i}`).textContent = '';
        row.querySelector(`.item${i}_Description1`).textContent = '';
        row.querySelector(`.itemprice${i}`).textContent = '';
        row.querySelector(`.itemBT${i}`).textContent = '';
    });

    // --- 填充静态数据 ---
    template.querySelector('.checkindate.rightdate').textContent = formatExcelDate(data['A']) || '';
    template.querySelector('.rightnowtime.checkintime').textContent = formatExcelTime(data['B']) || '';
    template.querySelector('.address1').textContent = data['E'] || '';
    template.querySelector('.address2').textContent = data['F'] || '';
    template.querySelector('.address3').textContent = data['G'] || '';

      function generateNumberGroups() {
        let groups = [];
        for (let i = 0; i < 3; i++) {
            // 生成 0~9999 的随机数，补足 4 位
            let group = Math.floor(Math.random() * 10000).toString().padStart(4, '0');
            groups.push(group);
        }
        return groups.join(' ');
    }


    template.querySelector('.val1').textContent = generateNumberGroups();
    template.querySelector('.val2').textContent = Math.floor(1000000 + Math.random() * 900000);



    // --- 动态处理商品和计算 ---
    let subtotal = 0;
    for (let i = 1; i <= 10; i++) {
        // 从第4个商品开始，Excel列有一个5列的偏移
        const offset = i >= 4 ? 5 : 0;

        const couponACol = XLSX.utils.encode_col(7 + (i - 1) * 5 + offset);  // H, M, R, AB...
        const couponBCol = XLSX.utils.encode_col(8 + (i - 1) * 5 + offset);  // I, N, S, AC...
        const priceCol = XLSX.utils.encode_col(9 + (i - 1) * 5 + offset);     // J, O, T, AD...
        const tagCol = XLSX.utils.encode_col(10 + (i - 1) * 5 + offset);   // K, P, U, AE...
        const descCol = XLSX.utils.encode_col(11 + (i - 1) * 5 + offset);     // L, Q, V, AF...

        const itemDesc = data[descCol];
        const itemPrice = parseFloat(data[priceCol]);

        if (itemDesc && !isNaN(itemPrice)) {
            const row = itemRows[i - 1];
            if (row) {
                row.style.display = 'table-row';
                row.querySelector(`.itemCouponA${i}`).textContent = data[couponBCol] || '';
                row.querySelector(`.itemCouponB${i}`).textContent = data[tagCol] || '';
                row.querySelector(`.item${i}_Description1`).textContent = data[couponACol] || '';
                row.querySelector(`.itemprice${i}`).textContent = itemPrice.toFixed(2);
                row.querySelector(`.itemBT${i}`).textContent = data[descCol] || 'N';
                subtotal += itemPrice;
            }
        }
    }

    // --- 计算和填充总计 ---
    const taxValue = parseFloat(data['BG']) || 0;
    const total = subtotal + taxValue;

    template.querySelector('.subtotal.comma').textContent = subtotal.toFixed(2);
    // template.querySelector('.taxname1').textContent = data['BF'] || 'TAX';
    template.querySelector('.taxprice1.comma').textContent = taxValue.toFixed(2);

    const totalElements = template.querySelectorAll('.total.comma');
    totalElements.forEach(el => el.textContent = total.toFixed(2));

    // --- 填充支付和消息 ---
    template.querySelector('.cardType').textContent = data['BH'] || '';
    template.querySelector('.lastNumber').textContent = data['BI'] || '';
    template.querySelector('.message1').textContent = data['BJ'] || '';

    // --- 处理Logo ---
    const logoImg = template.querySelector('.logo img');
    const logoFilename = data['BK'];
    console.log('data123,BK', data);

    if (logoFilename) {
        logoImg.src = `./logo/${logoFilename}`;
        logoImg.style.display = 'inline-block';
        logoImg.parentElement.parentElement.style.display = 'table-row';

    } else {
        logoImg.src = '';
        logoImg.style.display = 'none';
        logoImg.parentElement.parentElement.style.display = 'none';
    }

    // 触发高度调整
    if (typeof adjustElectronicStoreHeight === 'function') {
        adjustElectronicStoreHeight();
    }
}

// Online Furniture Shop模板数据填充函数
async function populateOnlineFurnitureShopReceipt(data) {
    const templateId = 'online-receipts-for-furniture-shop';
    const template = document.getElementById(templateId);
    if (!template) return;

    // --- 清理旧数据 ---
    const itemRows = template.querySelectorAll('.Pfare .remove_tr');
    itemRows.forEach((row, index) => {
        row.style.display = 'none';
        const i = index + 1;
        row.querySelector(`.qty${i}`).textContent = '';
        row.querySelector(`.farename${i}`).textContent = '';
        row.querySelector(`.farenumber${i}`).textContent = '';
        row.querySelector(`.fareprice${i}`).textContent = '';
        row.querySelector(`.fareN${i}`).textContent = '';
    });

    // --- 填充静态数据 ---
    // 处理日期和时间格式
    const dateValue = formatExcelDate(data['A']) || '09/11/2025';
    const enddateElement = template.querySelector('#online-receipts-for-furniture-shop .enddate');
    if (enddateElement) {
        enddateElement.textContent = dateValue;
    }

    const timeValue = formatExcelTime(data['B'], false) || '20:07';
    const endtimeElement = template.querySelector('#online-receipts-for-furniture-shop .endtime');
    if (endtimeElement) {
        endtimeElement.textContent = timeValue;
    }

    // SALE编号使用随机生成，不读取表格列
    const saleElement = template.querySelector('#online-receipts-for-furniture-shop .sale');
    if (saleElement) {
        // 生成20位随机数字
        let randomSale = '';
        for (let i = 0; i < 20; i++) {
            randomSale += Math.floor(Math.random() * 10);
        }
        saleElement.textContent = randomSale;
    }
    const address1Element = template.querySelector('#online-receipts-for-furniture-shop .address1');
    if (address1Element) {
        address1Element.textContent = data['E'] || '703-361-9900';
    }
    const address2Element = template.querySelector('#online-receipts-for-furniture-shop .address2');
    if (address2Element) {
        address2Element.textContent = data['F'] || '8346 shoppers Square';
    }
    const address3Element = template.querySelector('#online-receipts-for-furniture-shop .address3');
    if (address3Element) {
        address3Element.textContent = data['G'] || 'Manassas VA 201l1';
    }

    // --- 动态处理商品和计算 ---
    let subtotal = 0;
    let visibleItemCount = 0;

    for (let i = 1; i <= 10; i++) {
        // 正确的列映射：每个商品有独立的5列
        // 第1个商品：H,I,J,K,L (7,8,9,10,11)
        // 第2个商品：M,N,O,P,Q (12,13,14,15,16)
        // 第3个商品：R,S,T,U,V (17,18,19,20,21)
        // 以此类推...
        const nameCol = XLSX.utils.encode_col(7 + (i - 1) * 5);    // H, M, R, W, AB, AG, AL, AQ, AV, BA
        const numberCol = XLSX.utils.encode_col(8 + (i - 1) * 5);  // I, N, S, X, AC, AH, AM, AR, AW, BB
        const priceCol = XLSX.utils.encode_col(9 + (i - 1) * 5);   // J, O, T, Y, AD, AI, AN, AS, AX, BC
        const tagCol = XLSX.utils.encode_col(10 + (i - 1) * 5);    // K, P, U, Z, AE, AJ, AO, AT, AY, BD
        const qtyCol = XLSX.utils.encode_col(11 + (i - 1) * 5);    // L, Q, V, AA, AF, AK, AP, AU, AZ, BE

        const itemName = data[nameCol];
        const itemPrice = parseFloat(data[priceCol]);

        if (itemName && !isNaN(itemPrice)) {
            const row = itemRows[i - 1];
            if (row) {
                row.style.display = 'table-row';
                row.querySelector(`.qty${i}`).textContent = data[qtyCol] || '1';
                row.querySelector(`.farename${i}`).textContent = itemName;

                // 调试信息：输出商品编号相关数据
                console.log(`商品${i} - 列映射:`, {
                    nameCol, numberCol, priceCol, tagCol, qtyCol,
                    itemName: data[nameCol],
                    itemNumber: data[numberCol],
                    itemPrice: data[priceCol],
                    itemTag: data[tagCol],
                    itemQty: data[qtyCol]
                });

                // 特别检查第二个商品的N列数据
                if (i === 2) {
                    console.log('第二个商品详细调试:', {
                        'M列(商品名)': data['M'],
                        'N列(商品号)': data['N'],
                        'O列(价格)': data['O'],
                        'P列(标签)': data['P'],
                        'Q列(数量)': data['Q'],
                        'numberCol变量值': numberCol,
                        'data[numberCol]值': data[numberCol],
                        'N列是否存在': 'N' in data,
                        'N列值类型': typeof data['N'],
                        'N列值长度': data['N'] ? data['N'].length : 'undefined',
                        '原始data对象': data
                    });

                    // 检查N列前后的列是否有数据
                    console.log('N列前后列检查:', {
                        'L列': data['L'],
                        'M列': data['M'],
                        'N列': data['N'],
                        'O列': data['O'],
                        'P列': data['P'],
                        'Q列': data['Q'],
                        'R列': data['R']
                    });

                    // 检查所有列的数据
                    console.log('Excel原始数据检查:', {
                        'A列': data['A'], 'B列': data['B'], 'C列': data['C'], 'D列': data['D'], 'E列': data['E'],
                        'F列': data['F'], 'G列': data['G'], 'H列': data['H'], 'I列': data['I'], 'J列': data['J'],
                        'K列': data['K'], 'L列': data['L'], 'M列': data['M'], 'N列': data['N'], 'O列': data['O'],
                        'P列': data['P'], 'Q列': data['Q'], 'R列': data['R'], 'S列': data['S'], 'T列': data['T']
                    });
                }

                // 检查商品号是否为空或无效
                const itemNumber = data[numberCol];
                if (!itemNumber || itemNumber === 'N' || itemNumber === '') {
                    console.warn(`商品${i}的商品号无效:`, itemNumber, '列:', numberCol);
                }

                // 使用正确的列映射
                const farenumberElement = row.querySelector(`.farenumber${i}`);
                const farepriceElement = row.querySelector(`.fareprice${i}`);
                const fareNElement = row.querySelector(`.fareN${i}`);

                // 如果找不到元素，尝试其他选择器
                if (!farenumberElement && i === 2) {
                    console.log('尝试其他选择器查找farenumber2元素:', {
                        '直接查找': document.querySelector('.farenumber2'),
                        '在模板中查找': template.querySelector('.farenumber2'),
                        '在行中查找所有span': row.querySelectorAll('span')
                    });
                }

                // 特别调试第二个商品
                if (i === 2) {
                    console.log('第二个商品元素填充调试:', {
                        'farenumberElement存在': !!farenumberElement,
                        'data[numberCol]值': data[numberCol],
                        'numberCol': numberCol,
                        '填充前farenumber内容': farenumberElement ? farenumberElement.textContent : '元素不存在'
                    });
                }

                if (farenumberElement) {
                    farenumberElement.textContent = data[numberCol] || '';
                }
                if (farepriceElement) {
                    farepriceElement.textContent = itemPrice.toFixed(2);
                }
                if (fareNElement) {
                    fareNElement.textContent = ` ${data[tagCol] || 'N'}`;
                }

                // 再次检查第二个商品填充后的内容
                if (i === 2) {
                    console.log('第二个商品填充后内容:', {
                        'farenumber内容': farenumberElement ? farenumberElement.textContent : '元素不存在',
                        'fareprice内容': farepriceElement ? farepriceElement.textContent : '元素不存在',
                        'fareN内容': fareNElement ? fareNElement.textContent : '元素不存在'
                    });
                }
                subtotal += itemPrice;
                visibleItemCount++;
            }
        }
    }

    // --- 计算和填充总计 ---
    const taxValue = parseFloat(data['BH']) || 0;  // 修正：税费金额从BH列读取
    const total = subtotal + taxValue;

    template.querySelector('.subtotal').textContent = subtotal.toFixed(2);
    template.querySelector('.total').textContent = total.toFixed(2);
    // 更新所有显示总计的地方
    const totalElements = template.querySelectorAll('.total.comma');
    totalElements.forEach(el => el.textContent = total.toFixed(2));

    // --- 填充税费信息 ---
    template.querySelector('.taxname1').textContent = data['BF'] || 'TAX';  // 修正：TAX标签从BF列读取
    template.querySelector('.taxpro1').textContent = data['BG'] || '6.35';  // 修正：税费百分比从BG列读取
    template.querySelector('.taxprice1').textContent = taxValue.toFixed(2);

    // --- 填充支付和其他信息 ---
    template.querySelector('.swipe').textContent = data['BI'] || 'VISA';
    template.querySelector('.lastnumber').textContent = data['BJ'] || '8542';
    template.querySelector('.authoriz').textContent = data['BK'] || '582798';
    template.querySelector('.aid').textContent = data['BL'] || 'VR138L1F3V4T';
    template.querySelector('.message').textContent = data['BM'] || '✯✯✯ CUSTOMER COPY ✯✯✯';

    // --- 更新商品数量 ---
    updateOnlineFurnitureShopItemCount();

    // --- 处理Logo ---
    const logoImg = template.querySelector('.logo img');
    const logoFilename = data['BN'];
    console.log('data123,BN', data);

    if (logoFilename) {
        logoImg.src = `./logo/${logoFilename}`;
        logoImg.style.display = 'inline';
    } else {
        logoImg.style.display = 'none';
    }
}

// 更新第五个模板的商品总数
function updateOnlineFurnitureShopItemCount() {
    const templateId = '#online-receipts-for-furniture-shop';
    const template = document.querySelector(templateId);
    if (!template) return;

    // 计算显示的商品数量
    let itemCount = 0;
    const itemRows = template.querySelectorAll('.Pfare .remove_tr');

    itemRows.forEach(row => {
        if (row.style.display !== 'none') {
            itemCount++;
        }
    });

    // 更新商品总数显示
    const itemNumElement = template.querySelector('.itemnum');
    if (itemNumElement) {
        itemNumElement.textContent = itemCount;
    }

    console.log(`Online Furniture Shop模板商品总数已更新为: ${itemCount}`);
}

// ==================== 随机数字和条形码生成 ====================

// 检查 JsBarcode 库是否可用的辅助函数
function ensureJsBarcode(callback, delay = 500) {
    if (typeof JsBarcode !== 'undefined') {
        callback();
    } else {
        console.warn('JsBarcode 库未加载，延迟执行...');
        setTimeout(() => {
            if (typeof JsBarcode !== 'undefined') {
                callback();
            } else {
                console.error('JsBarcode 库加载失败，无法执行回调');
            }
        }, delay);
    }
}

document.addEventListener('DOMContentLoaded', function () {
    // 为第一个模板生成随机 REC/VCD
    function generateRandomNumber(length) {
        if (length <= 0) return '';
        let firstDigit = Math.floor(Math.random() * 9) + 1;
        if (length === 1) return firstDigit.toString();
        let restOfDigits = Array.from({ length: length - 1 }, () => Math.floor(Math.random() * 10)).join('');
        return firstDigit + restOfDigits;
    }

    const recVcdElement = document.querySelector('.recVcd');
    if (recVcdElement) {
        const recPart = `REC#${generateRandomNumber(1)}-${generateRandomNumber(4)}-${generateRandomNumber(4)}-${generateRandomNumber(4)}-${generateRandomNumber(4)}-${generateRandomNumber(1)}`;
        const vcdPart = `VCD#${generateRandomNumber(3)}-${generateRandomNumber(3)}-${generateRandomNumber(3)}`;
        recVcdElement.textContent = `${recPart} ${vcdPart}`;
    }
});

// 使用JsBarcode生成随机且有效的条形码
function applyRandomBarcode(svgElement) {
    if (!svgElement) {
        console.log("条形码元素不存在");
        return;
    }

    // 检查 JsBarcode 库是否可用
    if (typeof JsBarcode === 'undefined') {
        console.error("JsBarcode 库未定义，无法生成条形码");
        return;
    }

    let randomData = '';
    for (let i = 0; i < 22; i++) {
        randomData += Math.floor(Math.random() * 10);
    }

    console.log("正在生成条形码，数据:", randomData, "元素:", svgElement);

    try {
        // 检查是否是Electronic Store模板的条形码
        const isElectronicStore = svgElement.closest('#electronic-store-receipt-maker');

        if (isElectronicStore) {
            // Electronic Store模板：清空内容后重新生成条形码
            console.log("为Electronic Store模板生成条形码");

            // 清空现有内容
            svgElement.innerHTML = '';

            // 重置SVG样式，确保条形码可见
            svgElement.style.height = '70px';
            svgElement.style.width = '354px';
            svgElement.style.display = 'block';
            svgElement.style.background = '#ffffff';
            svgElement.style.border = 'none';
            svgElement.style.outline = 'none';
            svgElement.style.visibility = 'visible';
            svgElement.style.opacity = '1';

            // 设置正确的viewBox和尺寸（与源网站一致）
            svgElement.setAttribute('width', '354px');
            svgElement.setAttribute('height', '70px');
            svgElement.setAttribute('viewBox', '0 0 354 70');
            svgElement.setAttribute('x', '0px');
            svgElement.setAttribute('y', '0px');

            JsBarcode(svgElement, randomData, {
                format: "CODE128",
                width: 2.0,
                height: 50,
                displayValue: false,
                background: '#ffffff',
                lineColor: '#000000',
                margin: 10
            });

            // JsBarcode生成后，修复背景rect的宽度和高度
            const backgroundRect = svgElement.querySelector('rect[style*="fill:#ffffff"]');
            if (backgroundRect) {
                backgroundRect.setAttribute('width', '354');
                backgroundRect.setAttribute('height', '70');
            }

        } else {
            // 其他模板：清空内容后生成
            svgElement.innerHTML = '';

            // 重置SVG样式，确保条形码可见
            svgElement.style.height = '45px';
            svgElement.style.width = '354px';
            svgElement.style.display = 'block';
            svgElement.style.background = '#ffffff';
            svgElement.style.border = 'none';
            svgElement.style.outline = 'none';

            JsBarcode(svgElement, randomData, {
                format: "CODE128",
                width: 2.0,
                height: 40,
                displayValue: false,
                background: '#ffffff',
                lineColor: '#000000'
            });
        }

        console.log("条形码生成成功");
    } catch (e) {
        console.error("JsBarcode failed to generate:", e);
    }
}

// ==================== Grocery模板商品数量动态更新 ====================
// 动态计算并更新Grocery模板中的商品数量
function updateGroceryItemCount() {
    const groceryTemplate = document.getElementById('grocery-furniture-store-for-online-receipts');
    if (!groceryTemplate) return;

    // 查找所有显示的商品行（包含farename类的行，且不是隐藏的）
    const productTable = groceryTemplate.querySelector('.product-fare-table');
    if (!productTable) return;

    // 计算可见的商品项目数量
    const visibleItems = productTable.querySelectorAll('.remove_tr');
    let itemCount = 0;

    visibleItems.forEach(item => {
        // 检查该行是否可见且包含商品信息
        const farenames = item.querySelectorAll('[class*="farename"]');
        const hasContent = Array.from(farenames).some(span => {
            const text = span.textContent.trim();
            return text && text !== '';
        });

        if (hasContent && item.style.display !== 'none') {
            itemCount++;
        }
    });

    // 更新ITEMS SOLD显示的数字
    const itemNumElement = groceryTemplate.querySelector('.itemnum');
    if (itemNumElement) {
        itemNumElement.textContent = itemCount;
        console.log(`Grocery模板商品数量已更新为: ${itemCount}`);
    }
}

// 页面加载时初始化计算
document.addEventListener('DOMContentLoaded', function () {
    // 延迟执行确保DOM完全加载
    setTimeout(updateGroceryItemCount, 100);
});

// 提供全局函数供外部调用（例如批量生成时）
window.updateGroceryItemCount = updateGroceryItemCount;

// ==================== Logo上传功能 ====================
let currentUploadedLogo = null;

document.addEventListener('DOMContentLoaded', function () {
    // Logo上传功能
    const uploadLogoBtn = document.getElementById('upload-logo-btn');
    if (uploadLogoBtn) {
        uploadLogoBtn.addEventListener('click', async function () {
            const fileInput = document.getElementById('logo-file-input');
            const files = fileInput.files;

            if (files.length === 0) {
                alert('请先选择一个或多个图片文件！');
                return;
            }

            const formData = new FormData();
            for (let i = 0; i < files.length; i++) {
                const file = files[i];
                // 检查文件类型
                if (!file.type.startsWith('image/')) {
                    alert(`文件 "${file.name}" 不是有效的图片文件，将被忽略。`);
                    continue; // 跳过非图片文件
                }
                formData.append('logos', file);
            }

            // 检查是否有有效文件被添加
            if (!formData.has('logos')) {
                alert('没有选择有效的图片文件进行上传。');
                return;
            }

            try {
                const response = await fetch('/upload-logo', {
                    method: 'POST',
                    body: formData
                });

                if (!response.ok) {
                    const errorData = await response.json();
                    throw new Error(errorData.error || '服务器响应错误');
                }

                const result = await response.json();

                if (result.filePaths && result.filePaths.length > 0) {
                    // 使用上传的第一个logo来更新显示
                    const firstLogoSrc = result.filePaths[0];
                    currentUploadedLogo = firstLogoSrc;

                    // 更新预览
                    updateLogoPreview(firstLogoSrc);

                    // 更新所有模板中的logo并等待加载完成
                    await updateTemplateLogos(firstLogoSrc);

                    // 调整Electronic Store模板高度
                    if (window.adjustElectronicStoreHeight) {
                        setTimeout(() => {
                            window.adjustElectronicStoreHeight();
                        }, 100);
                    }

                    alert(`${result.filePaths.length}个Logo上传成功！`);
                } else {
                    alert('Logo上传失败，服务器未返回文件路径。');
                }

            } catch (error) {
                console.error('Logo上传失败:', error);
                alert(`Logo上传失败: ${error.message}`);
            }
        });
    }

    // 移除Logo功能
    const removeLogoBtn = document.getElementById('remove-logo-btn');
    if (removeLogoBtn) {
        removeLogoBtn.addEventListener('click', async function () {
            currentUploadedLogo = null;

            // 清除预览
            const preview = document.getElementById('logo-preview');
            if (preview) {
                preview.innerHTML = '<span style="color: #666;">选择图片文件来预览logo</span>';
            }

            // 清除模板中的logo
            await updateTemplateLogos('');

            // 调整Electronic Store模板高度
            if (window.adjustElectronicStoreHeight) {
                setTimeout(() => {
                    window.adjustElectronicStoreHeight();
                }, 100);
            }

            // 清除文件输入
            const fileInput = document.getElementById('logo-file-input');
            if (fileInput) {
                fileInput.value = '';
            }

            alert('Logo已移除！');
        });
    }

    // 文件选择变化时的预览
    const logoFileInput = document.getElementById('logo-file-input');
    if (logoFileInput) {
        logoFileInput.addEventListener('change', function (e) {
            const file = e.target.files[0]; // 只预览第一个选择的文件
            if (file && file.type.startsWith('image/')) {
                const reader = new FileReader();
                reader.onload = function (e) {
                    updateLogoPreview(e.target.result);
                };
                reader.readAsDataURL(file);
            }
        });
    }

    // 页面加载时恢复上传的logo（如果有的话）
    if (currentUploadedLogo) {
        updateTemplateLogos(currentUploadedLogo);
        updateLogoPreview(currentUploadedLogo);
    }
});

// 更新logo预览
function updateLogoPreview(logoSrc) {
    const preview = document.getElementById('logo-preview');
    if (preview) {
        preview.innerHTML = `<img src="${logoSrc}" style="max-width: 100%; max-height: 60px; object-fit: contain;" alt="Logo Preview">`;
    }
}

// 更新所有模板中的logo
async function updateTemplateLogos(logoSrc) {
    const allTemplates = document.querySelectorAll('.receipt-template');
    const promises = [];
    allTemplates.forEach(template => {
        const logoImg = template.querySelector('.logo img');
        if (logoImg) {
            logoImg.src = logoSrc;

            if (logoSrc) {
                const promise = new Promise((resolve) => {
                    logoImg.onload = resolve;
                    logoImg.onerror = () => {
                        console.error("Logo failed to load");
                        resolve(); // 即使失败也继续，避免阻塞
                    };
                });
                promises.push(promise);
            }
        }
    });
    await Promise.all(promises);
}

// ==================== Grocery模板随机数字生成 ====================
// 生成指定位数的随机数字
function generateRandomNumber(length) {
    if (length <= 0) return '';
    let result = '';
    // 第一位不能为0（除非只有1位）
    if (length === 1) {
        result = Math.floor(Math.random() * 10).toString();
    } else {
        result = Math.floor(Math.random() * 9 + 1).toString(); // 1-9
        for (let i = 1; i < length; i++) {
            result += Math.floor(Math.random() * 10).toString(); // 0-9
        }
    }
    return result;
}

// 为Grocery模板生成随机数字
function generateGroceryRandomNumbers() {
    const groceryTemplate = document.getElementById('grocery-furniture-store-for-online-receipts');
    if (!groceryTemplate) return;

    // 生成ST#数字（4位）
    const stElement = groceryTemplate.querySelector('.st');
    if (stElement) {
        stElement.textContent = generateRandomNumber(4);
    }

    // 生成OP#数字（8位）
    const opElement = groceryTemplate.querySelector('.op');
    if (opElement) {
        opElement.textContent = generateRandomNumber(8);
    }

    // 生成TE#数字（2位）
    const teElement = groceryTemplate.querySelector('.te');
    if (teElement) {
        teElement.textContent = generateRandomNumber(2);
    }

    // 生成TR#数字（5位）
    const trElement = groceryTemplate.querySelector('.tr');
    if (trElement) {
        trElement.textContent = generateRandomNumber(5);
    }

    // 生成REF#数字（12位）
    const refElement = groceryTemplate.querySelector('.ref');
    if (refElement) {
        refElement.textContent = generateRandomNumber(12);
    }

    // 生成TERMINAL#数字（10位）
    const terminalElement = groceryTemplate.querySelector('.terminal');
    if (terminalElement) {
        terminalElement.textContent = generateRandomNumber(10);
    }

    console.log('Grocery模板随机数字已更新');
}

// 页面加载时生成随机数字
document.addEventListener('DOMContentLoaded', function () {
    // 延迟执行确保所有元素都已加载
    setTimeout(generateGroceryRandomNumbers, 200);
});

// 提供全局函数供外部调用
window.generateGroceryRandomNumbers = generateGroceryRandomNumbers;

// ==================== Electronic Store模板随机数字生成 ====================
// 为Electronic Store模板生成随机数字
function generateElectronicStoreRandomNumbers() {
    const electronicTemplate = document.getElementById('electronic-store-receipt-maker');
    if (!electronicTemplate) return;

    // 生成VAL数字（格式：XXXX-XXXX-XXXX-XXXX）
    const valElement = electronicTemplate.querySelector('.val');
    if (valElement) {
        const val1 = generateRandomNumber(4);
        const val2 = generateRandomNumber(4);
        const val3 = generateRandomNumber(4);
        const val4 = generateRandomNumber(4);
        valElement.textContent = `${val1}-${val2}-${val3}-${val4}`;
    }

    // 生成VAL1数字（格式：XXXX XXXX XXXX）
    const val1Element = electronicTemplate.querySelector('.val1');
    if (val1Element) {
        const val1a = generateRandomNumber(4);
        const val1b = generateRandomNumber(4);
        const val1c = generateRandomNumber(4);
        val1Element.textContent = `${val1a} ${val1b} ${val1c}`;
    }

    // 生成VAL2数字（7位数字）
    const val2Element = electronicTemplate.querySelector('.val2');
    if (val2Element) {
        val2Element.textContent = generateRandomNumber(7);
    }

    // 生成APPROVAL数字（6位数字）
    const approvalElement = electronicTemplate.querySelector('.approval');
    if (approvalElement) {
        approvalElement.textContent = generateRandomNumber(6);
    }

    // 生成REFERENCE NUMBER（格式：XXXXX XX XXX X）
    const refNoElement = electronicTemplate.querySelector('.refNo');
    if (refNoElement) {
        const ref1 = generateRandomNumber(5);
        const ref2 = generateRandomNumber(2);
        const ref3 = generateRandomNumber(3);
        const ref4 = generateRandomNumber(1);
        refNoElement.textContent = `${ref1} ${ref2} ${ref3} ${ref4}`;
    }

    console.log('Electronic Store模板随机数字已更新');
}

// 页面加载时生成Electronic Store随机数字
document.addEventListener('DOMContentLoaded', function () {
    // 延迟执行确保所有元素都已加载
    setTimeout(generateElectronicStoreRandomNumbers, 200);
});

// 提供全局函数供外部调用
window.generateElectronicStoreRandomNumbers = generateElectronicStoreRandomNumbers;

// ==================== Electronic Store模板动态高度调整 ====================
// 动态调整Electronic Store模板的margin-bottom
function adjustElectronicStoreHeight() {
    const electronicTemplate = document.getElementById('electronic-store-receipt-maker');
    if (!electronicTemplate) return;

    const formGroup = electronicTemplate.querySelector('.form-group.mb-1');
    if (!formGroup) return;

    // 计算内容高度
    const contentHeight = formGroup.scrollHeight;
    const logoImg = electronicTemplate.querySelector('.logo img');
    const hasLogo = logoImg && logoImg.src && logoImg.src.trim() !== '';

    // 根据内容高度和是否有Logo来动态设置margin-bottom
    let marginBottom = 0;

    if (hasLogo) {
        // 有Logo时的基础margin
        marginBottom = 0.2; // 原为 0.5

        // 根据内容高度调整
        if (contentHeight > 2000) {
            marginBottom = 0.3; // 原为 0.8
        } else if (contentHeight > 1500) {
            marginBottom = 0.4; // 原为 1.0
        } else if (contentHeight < 800) {
            marginBottom = 0.5; // 原为 1.2
        }
    } else {
        // 无Logo时根据内容高度调整
        if (contentHeight > 2000) {
            marginBottom = 0.1; // 原为 0.2
        } else if (contentHeight > 1500) {
            marginBottom = 0.2; // 原为 0.5
        } else if (contentHeight < 800) {
            marginBottom = 0.3; // 原为 0.8
        }
    }

    // 应用动态margin-bottom，使用!important确保优先级
    formGroup.style.setProperty('margin-bottom', `${marginBottom}rem`, 'important');

    console.log(`Electronic Store高度调整: 内容高度=${contentHeight}px, 有Logo=${hasLogo}, margin-bottom=${marginBottom}rem`);
}

// 页面加载时调整高度
document.addEventListener('DOMContentLoaded', function () {
    setTimeout(adjustElectronicStoreHeight, 300);

    // 等待 JsBarcode 库加载完成
    function waitForJsBarcode() {
        if (typeof JsBarcode !== 'undefined') {
            console.log("JsBarcode库已加载");
            // 为所有模板生成条形码
            generateAllBarcodes();
        } else {
            console.log("等待 JsBarcode 库加载...");
            setTimeout(waitForJsBarcode, 100);
        }
    }

    // 开始等待 JsBarcode 库
    waitForJsBarcode();

    // 为特定模板生成随机数（不依赖 JsBarcode）
    generateGroceryRandomNumbers();
    generateElectronicStoreRandomNumbers();
    generateGeneralGroceryRandomNumbers();
    generateOnlineFurnitureShopRandomNumbers();
});

// 生成所有模板的条形码
function generateAllBarcodes() {
    console.log("开始生成所有模板的条形码");

    // 为Retail模板生成条形码
    const retailBarcode = document.querySelector('#retail-electronic-home-Improvements-chain-store .barcode');
    console.log("Retail条形码元素:", retailBarcode);
    if (retailBarcode) {
        applyRandomBarcode(retailBarcode);
    }

    // 为Grocery模板生成条形码
    const groceryBarcode = document.querySelector('#grocery-furniture-store-for-online-receipts .barcode');
    console.log("Grocery条形码元素:", groceryBarcode);
    if (groceryBarcode) {
        applyRandomBarcode(groceryBarcode);
    }

    // 为Electronic Store模板生成条形码
    const electronicBarcode = document.querySelector('#electronic-store-receipt-maker .barcode');
    console.log("Electronic Store条形码元素:", electronicBarcode);
    if (electronicBarcode) {
        applyRandomBarcode(electronicBarcode);
    }

    // 为Online Furniture Shop模板生成条形码
    const furnitureBarcode = document.querySelector('#online-receipts-for-furniture-shop .barcode');
    console.log("Online Furniture Shop条形码元素:", furnitureBarcode);
    if (furnitureBarcode) {
        applyRandomBarcode(furnitureBarcode);
    }
}

// 手动触发条形码生成（用于调试）
function manualGenerateBarcodes() {
    console.log("手动生成条形码");

    // 为Retail模板生成条形码
    const retailBarcode = document.querySelector('#retail-electronic-home-Improvements-chain-store .barcode');
    if (retailBarcode) {
        console.log("为Retail模板生成条形码");
        applyRandomBarcode(retailBarcode);
    }

    // 为Electronic Store模板生成条形码
    const electronicBarcode = document.querySelector('#electronic-store-receipt-maker .barcode');
    if (electronicBarcode) {
        console.log("为Electronic Store模板生成条形码");
        applyRandomBarcode(electronicBarcode);
    }
}

// 将函数暴露到全局作用域，方便调试
window.manualGenerateBarcodes = manualGenerateBarcodes;

// 专门为Electronic Store模板生成条形码的测试函数
function testElectronicStoreBarcode() {
    console.log("测试Electronic Store条形码生成");
    const electronicBarcode = document.querySelector('#electronic-store-receipt-maker .barcode');
    if (electronicBarcode) {
        console.log("找到Electronic Store条形码元素:", electronicBarcode);

        // 使用统一的条形码生成函数
        applyRandomBarcode(electronicBarcode);
    } else {
        console.log("未找到Electronic Store条形码元素");
    }
}

window.testElectronicStoreBarcode = testElectronicStoreBarcode;

// 检查Electronic Store条形码内容的函数
function checkElectronicStoreBarcode() {
    const electronicBarcode = document.querySelector('#electronic-store-receipt-maker .barcode');
    if (electronicBarcode) {
        console.log("Electronic Store条形码元素:", electronicBarcode);
        console.log("SVG内容:", electronicBarcode.innerHTML);
        console.log("SVG子元素数量:", electronicBarcode.children.length);
        console.log("SVG属性:", {
            width: electronicBarcode.getAttribute('width'),
            height: electronicBarcode.getAttribute('height'),
            viewBox: electronicBarcode.getAttribute('viewBox')
        });

        // 检查是否有rect元素
        const rects = electronicBarcode.querySelectorAll('rect');
        console.log("rect元素数量:", rects.length);
        if (rects.length > 0) {
            console.log("第一个rect元素:", rects[0]);
        }
    } else {
        console.log("未找到Electronic Store条形码元素");
    }
}

window.checkElectronicStoreBarcode = checkElectronicStoreBarcode;

// 强制显示Electronic Store条形码的函数
function forceShowElectronicStoreBarcode() {
    const electronicBarcode = document.querySelector('#electronic-store-receipt-maker .barcode');
    if (electronicBarcode) {
        console.log("强制显示Electronic Store条形码");

        // 强制设置所有样式
        electronicBarcode.style.cssText = `
            display: block !important;
            width: 354px !important;
            height: 70px !important;
            background: #ffffff !important;
            border: 1px solid #000 !important;
            visibility: visible !important;
            opacity: 1 !important;
            margin: 10px auto !important;
            position: relative !important;
            z-index: 9999 !important;
        `;

        // 强制设置SVG样式
        const svg = electronicBarcode.querySelector('svg');
        if (svg) {
            svg.style.cssText = `
                display: block !important;
                width: 354px !important;
                height: 70px !important;
                background: #ffffff !important;
                visibility: visible !important;
                opacity: 1 !important;
            `;
        }

        // 强制设置所有rect元素样式
        const rects = electronicBarcode.querySelectorAll('rect');
        rects.forEach(rect => {
            rect.style.cssText = `
                visibility: visible !important;
                opacity: 1 !important;
                fill: #000000 !important;
            `;
        });

        console.log("强制显示完成，rect元素数量:", rects.length);
    } else {
        console.log("未找到Electronic Store条形码元素");
    }
}

window.forceShowElectronicStoreBarcode = forceShowElectronicStoreBarcode;

// 直接替换SVG内容来测试条形码显示
function replaceElectronicStoreBarcode() {
    const electronicBarcode = document.querySelector('#electronic-store-receipt-maker .barcode');
    if (electronicBarcode) {
        console.log("直接替换Electronic Store条形码内容");

        // 直接设置SVG内容
        electronicBarcode.innerHTML = `
            <rect x="0" y="0" width="354" height="70" style="fill:#ffffff;"></rect>
            <g transform="translate(10, 10)" style="fill:#000000;">
                <rect x="0" y="0" width="4" height="50"></rect>
                <rect x="6" y="0" width="2" height="50"></rect>
                <rect x="12" y="0" width="6" height="50"></rect>
                <rect x="22" y="0" width="4" height="50"></rect>
                <rect x="28" y="0" width="6" height="50"></rect>
                <rect x="38" y="0" width="2" height="50"></rect>
                <rect x="44" y="0" width="4" height="50"></rect>
                <rect x="56" y="0" width="2" height="50"></rect>
                <rect x="62" y="0" width="2" height="50"></rect>
                <rect x="66" y="0" width="2" height="50"></rect>
                <rect x="74" y="0" width="4" height="50"></rect>
                <rect x="80" y="0" width="2" height="50"></rect>
                <rect x="88" y="0" width="6" height="50"></rect>
                <rect x="96" y="0" width="2" height="50"></rect>
                <rect x="104" y="0" width="4" height="50"></rect>
                <rect x="110" y="0" width="4" height="50"></rect>
                <rect x="122" y="0" width="2" height="50"></rect>
                <rect x="126" y="0" width="2" height="50"></rect>
                <rect x="132" y="0" width="2" height="50"></rect>
                <rect x="136" y="0" width="2" height="50"></rect>
                <rect x="142" y="0" width="8" height="50"></rect>
                <rect x="154" y="0" width="4" height="50"></rect>
                <rect x="162" y="0" width="6" height="50"></rect>
                <rect x="170" y="0" width="2" height="50"></rect>
                <rect x="176" y="0" width="6" height="50"></rect>
                <rect x="186" y="0" width="4" height="50"></rect>
                <rect x="192" y="0" width="2" height="50"></rect>
                <rect x="198" y="0" width="4" height="50"></rect>
                <rect x="204" y="0" width="4" height="50"></rect>
                <rect x="212" y="0" width="4" height="50"></rect>
                <rect x="220" y="0" width="4" height="50"></rect>
                <rect x="226" y="0" width="4" height="50"></rect>
                <rect x="234" y="0" width="4" height="50"></rect>
                <rect x="242" y="0" width="6" height="50"></rect>
                <rect x="250" y="0" width="2" height="50"></rect>
                <rect x="254" y="0" width="8" height="50"></rect>
                <rect x="264" y="0" width="2" height="50"></rect>
                <rect x="270" y="0" width="6" height="50"></rect>
                <rect x="278" y="0" width="4" height="50"></rect>
                <rect x="286" y="0" width="2" height="50"></rect>
                <rect x="296" y="0" width="2" height="50"></rect>
                <rect x="302" y="0" width="4" height="50"></rect>
                <rect x="308" y="0" width="4" height="50"></rect>
                <rect x="318" y="0" width="6" height="50"></rect>
                <rect x="326" y="0" width="2" height="50"></rect>
                <rect x="330" y="0" width="4" height="50"></rect>
            </g>
        `;

        // 设置正确的属性
        electronicBarcode.setAttribute('width', '354px');
        electronicBarcode.setAttribute('height', '70px');
        electronicBarcode.setAttribute('viewBox', '0 0 354 70');
        electronicBarcode.style.cssText = `
            display: block !important;
            width: 354px !important;
            height: 70px !important;
            background: #ffffff !important;
            border: 2px solid red !important;
            visibility: visible !important;
            opacity: 1 !important;
            margin: 10px auto !important;
        `;

        console.log("直接替换完成，应该能看到红色边框的条形码");
    } else {
        console.log("未找到Electronic Store条形码元素");
    }
}

window.replaceElectronicStoreBarcode = replaceElectronicStoreBarcode;

// 修复现有条形码的显示问题
function fixElectronicStoreBarcode() {
    const electronicBarcode = document.querySelector('#electronic-store-receipt-maker .barcode');
    if (electronicBarcode) {
        console.log("修复Electronic Store条形码显示");

        // 修复背景rect的尺寸
        const backgroundRect = electronicBarcode.querySelector('rect[style*="fill:#ffffff"]');
        if (backgroundRect) {
            backgroundRect.setAttribute('width', '354');
            backgroundRect.setAttribute('height', '70');
            backgroundRect.style.width = '354px';
            backgroundRect.style.height = '70px';
            console.log("背景rect已修复");
        }

        // 确保SVG属性正确
        electronicBarcode.setAttribute('width', '354px');
        electronicBarcode.setAttribute('height', '70px');
        electronicBarcode.setAttribute('viewBox', '0 0 354 70');
        electronicBarcode.style.width = '354px';
        electronicBarcode.style.height = '70px';

        // 强制设置样式确保可见
        electronicBarcode.style.cssText = `
            display: block !important;
            width: 354px !important;
            height: 70px !important;
            background: #ffffff !important;
            border: 2px solid blue !important;
            visibility: visible !important;
            opacity: 1 !important;
            margin: 10px auto !important;
        `;

        // 确保所有rect元素可见
        const rects = electronicBarcode.querySelectorAll('rect');
        rects.forEach(rect => {
            rect.style.cssText = `
                visibility: visible !important;
                opacity: 1 !important;
                fill: #000000 !important;
            `;
        });

        console.log("修复完成，应该能看到蓝色边框的条形码，rect元素数量:", rects.length);
    } else {
        console.log("未找到Electronic Store条形码元素");
    }
}

window.fixElectronicStoreBarcode = fixElectronicStoreBarcode;

// 测试第三个模板条形码显示的完整函数
function testElectronicStoreBarcodeDisplay() {
    console.log("开始测试第三个模板条形码显示");

    const electronicBarcode = document.querySelector('#electronic-store-receipt-maker .barcode');
    if (!electronicBarcode) {
        console.error("未找到第三个模板的条形码元素");
        return;
    }

    console.log("找到条形码元素:", electronicBarcode);
    console.log("元素标签名:", electronicBarcode.tagName);

    // 检查当前样式
    const computedStyle = window.getComputedStyle(electronicBarcode);
    console.log("当前样式:", {
        display: computedStyle.display,
        width: computedStyle.width,
        height: computedStyle.height,
        visibility: computedStyle.visibility,
        opacity: computedStyle.opacity
    });

    // 强制应用正确的样式
    electronicBarcode.style.cssText = `
        display: block !important;
        width: 354px !important;
        height: 70px !important;
        background: #ffffff !important;
        visibility: visible !important;
        opacity: 1 !important;
        margin: 0 auto !important;
    `;

    // 检查是否是SVG元素
    if (electronicBarcode.tagName === 'SVG') {
        console.log("条形码元素本身就是SVG");

        // 检查rect元素
        const rects = electronicBarcode.querySelectorAll('rect');
        console.log("SVG中的rect元素数量:", rects.length);

        if (rects.length > 0) {
            console.log("条形码rect元素存在，应该可见");
            rects.forEach((rect, index) => {
                if (index < 5) { // 只显示前5个rect的信息
                    console.log(`Rect ${index}:`, {
                        x: rect.getAttribute('x'),
                        y: rect.getAttribute('y'),
                        width: rect.getAttribute('width'),
                        height: rect.getAttribute('height'),
                        fill: rect.getAttribute('fill'),
                        style: rect.getAttribute('style')
                    });
                }
            });
        } else {
            console.log("SVG中没有rect元素，需要生成条形码");
            // 生成条形码
            applyRandomBarcode(electronicBarcode);
        }
    } else {
        // 如果不是SVG，查找SVG子元素
        const svgElement = electronicBarcode.querySelector('svg');
        if (svgElement) {
            console.log("找到SVG子元素:", svgElement);
            svgElement.style.cssText = `
                display: block !important;
                width: 354px !important;
                height: 70px !important;
                background: #ffffff !important;
                visibility: visible !important;
                opacity: 1 !important;
            `;

            // 检查rect元素
            const rects = svgElement.querySelectorAll('rect');
            console.log("SVG中的rect元素数量:", rects.length);

            rects.forEach((rect, index) => {
                if (index < 5) { // 只显示前5个rect的信息
                    console.log(`Rect ${index}:`, {
                        x: rect.getAttribute('x'),
                        y: rect.getAttribute('y'),
                        width: rect.getAttribute('width'),
                        height: rect.getAttribute('height'),
                        fill: rect.getAttribute('fill')
                    });
                }
            });
        } else {
            console.error("未找到SVG元素");
        }
    }

    console.log("条形码显示测试完成");
}

window.testElectronicStoreBarcodeDisplay = testElectronicStoreBarcodeDisplay;

// 强制为第三个模板生成条形码
function forceGenerateElectronicStoreBarcode() {
    console.log("强制为第三个模板生成条形码");

    const electronicBarcode = document.querySelector('#electronic-store-receipt-maker .barcode');
    if (!electronicBarcode) {
        console.error("未找到第三个模板的条形码元素");
        return;
    }

    console.log("找到条形码元素，开始生成...");

    // 直接调用条形码生成函数
    applyRandomBarcode(electronicBarcode);

    // 等待一下再检查结果
    setTimeout(() => {
        const rects = electronicBarcode.querySelectorAll('rect');
        console.log("生成后SVG中的rect元素数量:", rects.length);

        if (rects.length > 0) {
            console.log("条形码生成成功！");
            // 强制设置样式确保可见
            electronicBarcode.style.cssText = `
                display: block !important;
                width: 354px !important;
                height: 70px !important;
                background: #ffffff !important;
                visibility: visible !important;
                opacity: 1 !important;
                margin: 0 auto !important;
                border: 2px solid red !important;
            `;
        } else {
            console.error("条形码生成失败，没有rect元素");
        }
    }, 100);
}

window.forceGenerateElectronicStoreBarcode = forceGenerateElectronicStoreBarcode;

// 详细调试条形码内容
function debugElectronicStoreBarcode() {
    const electronicBarcode = document.querySelector('#electronic-store-receipt-maker .barcode');
    if (electronicBarcode) {
        console.log("=== 详细调试Electronic Store条形码 ===");
        console.log("SVG元素:", electronicBarcode);
        console.log("SVG属性:", {
            width: electronicBarcode.getAttribute('width'),
            height: electronicBarcode.getAttribute('height'),
            viewBox: electronicBarcode.getAttribute('viewBox')
        });

        // 检查所有子元素
        console.log("SVG子元素数量:", electronicBarcode.children.length);
        Array.from(electronicBarcode.children).forEach((child, index) => {
            console.log(`子元素 ${index}:`, child.tagName, child);
            if (child.tagName === 'g') {
                console.log(`  - g元素transform:`, child.getAttribute('transform'));
                console.log(`  - g元素子元素数量:`, child.children.length);
                Array.from(child.children).forEach((rect, rectIndex) => {
                    if (rectIndex < 5) { // 只显示前5个rect
                        console.log(`    rect ${rectIndex}:`, {
                            x: rect.getAttribute('x'),
                            y: rect.getAttribute('y'),
                            width: rect.getAttribute('width'),
                            height: rect.getAttribute('height'),
                            style: rect.getAttribute('style')
                        });
                    }
                });
            }
        });

        // 检查rect元素的位置
        const rects = electronicBarcode.querySelectorAll('rect');
        console.log("总rect元素数量:", rects.length);
        if (rects.length > 0) {
            const firstRect = rects[0];
            const lastRect = rects[rects.length - 1];
            console.log("第一个rect:", {
                x: firstRect.getAttribute('x'),
                y: firstRect.getAttribute('y'),
                width: firstRect.getAttribute('width'),
                height: firstRect.getAttribute('height')
            });
            console.log("最后一个rect:", {
                x: lastRect.getAttribute('x'),
                y: lastRect.getAttribute('y'),
                width: lastRect.getAttribute('width'),
                height: lastRect.getAttribute('height')
            });
        }

        // 检查是否有超出viewBox的rect
        const viewBoxWidth = 354;
        const viewBoxHeight = 70;
        let outOfBounds = 0;
        rects.forEach(rect => {
            const x = parseFloat(rect.getAttribute('x')) || 0;
            const y = parseFloat(rect.getAttribute('y')) || 0;
            const width = parseFloat(rect.getAttribute('width')) || 0;
            const height = parseFloat(rect.getAttribute('height')) || 0;

            if (x + width > viewBoxWidth || y + height > viewBoxHeight) {
                outOfBounds++;
            }
        });
        console.log("超出viewBox的rect数量:", outOfBounds);

    } else {
        console.log("未找到Electronic Store条形码元素");
    }
}

window.debugElectronicStoreBarcode = debugElectronicStoreBarcode;

// 创建简单的测试条形码
function createSimpleBarcode() {
    const electronicBarcode = document.querySelector('#electronic-store-receipt-maker .barcode');
    if (electronicBarcode) {
        console.log("创建简单的测试条形码");

        // 清空并创建简单的条形码
        electronicBarcode.innerHTML = `
            <rect x="0" y="0" width="354" height="70" fill="#ffffff" stroke="#000000" stroke-width="1"></rect>
            <rect x="10" y="10" width="4" height="50" fill="#000000"></rect>
            <rect x="16" y="10" width="2" height="50" fill="#000000"></rect>
            <rect x="20" y="10" width="6" height="50" fill="#000000"></rect>
            <rect x="28" y="10" width="4" height="50" fill="#000000"></rect>
            <rect x="34" y="10" width="6" height="50" fill="#000000"></rect>
            <rect x="42" y="10" width="2" height="50" fill="#000000"></rect>
            <rect x="46" y="10" width="4" height="50" fill="#000000"></rect>
            <rect x="58" y="10" width="2" height="50" fill="#000000"></rect>
            <rect x="64" y="10" width="2" height="50" fill="#000000"></rect>
            <rect x="68" y="10" width="2" height="50" fill="#000000"></rect>
            <rect x="76" y="10" width="4" height="50" fill="#000000"></rect>
            <rect x="82" y="10" width="2" height="50" fill="#000000"></rect>
            <rect x="90" y="10" width="6" height="50" fill="#000000"></rect>
            <rect x="98" y="10" width="2" height="50" fill="#000000"></rect>
            <rect x="106" y="10" width="4" height="50" fill="#000000"></rect>
            <rect x="112" y="10" width="4" height="50" fill="#000000"></rect>
            <rect x="124" y="10" width="2" height="50" fill="#000000"></rect>
            <rect x="128" y="10" width="2" height="50" fill="#000000"></rect>
            <rect x="134" y="10" width="2" height="50" fill="#000000"></rect>
            <rect x="138" y="10" width="2" height="50" fill="#000000"></rect>
            <rect x="144" y="10" width="8" height="50" fill="#000000"></rect>
            <rect x="156" y="10" width="4" height="50" fill="#000000"></rect>
            <rect x="164" y="10" width="6" height="50" fill="#000000"></rect>
            <rect x="172" y="10" width="2" height="50" fill="#000000"></rect>
            <rect x="178" y="10" width="6" height="50" fill="#000000"></rect>
            <rect x="188" y="10" width="4" height="50" fill="#000000"></rect>
            <rect x="194" y="10" width="2" height="50" fill="#000000"></rect>
            <rect x="200" y="10" width="4" height="50" fill="#000000"></rect>
            <rect x="206" y="10" width="4" height="50" fill="#000000"></rect>
            <rect x="214" y="10" width="4" height="50" fill="#000000"></rect>
            <rect x="222" y="10" width="4" height="50" fill="#000000"></rect>
            <rect x="228" y="10" width="4" height="50" fill="#000000"></rect>
            <rect x="236" y="10" width="4" height="50" fill="#000000"></rect>
            <rect x="244" y="10" width="6" height="50" fill="#000000"></rect>
            <rect x="252" y="10" width="2" height="50" fill="#000000"></rect>
            <rect x="256" y="10" width="8" height="50" fill="#000000"></rect>
            <rect x="266" y="10" width="2" height="50" fill="#000000"></rect>
            <rect x="272" y="10" width="6" height="50" fill="#000000"></rect>
            <rect x="280" y="10" width="4" height="50" fill="#000000"></rect>
            <rect x="288" y="10" width="2" height="50" fill="#000000"></rect>
            <rect x="298" y="10" width="2" height="50" fill="#000000"></rect>
            <rect x="304" y="10" width="4" height="50" fill="#000000"></rect>
            <rect x="312" y="10" width="4" height="50" fill="#000000"></rect>
            <rect x="322" y="10" width="6" height="50" fill="#000000"></rect>
            <rect x="330" y="10" width="2" height="50" fill="#000000"></rect>
            <rect x="334" y="10" width="4" height="50" fill="#000000"></rect>
        `;

        // 设置正确的属性
        electronicBarcode.setAttribute('width', '354px');
        electronicBarcode.setAttribute('height', '70px');
        electronicBarcode.setAttribute('viewBox', '0 0 354 70');
        electronicBarcode.style.cssText = `
            display: block !important;
            width: 354px !important;
            height: 70px !important;
            background: #ffffff !important;
            border: 3px solid green !important;
            visibility: visible !important;
            opacity: 1 !important;
            margin: 10px auto !important;
        `;

        console.log("简单条形码创建完成，应该能看到绿色边框的条形码");
    } else {
        console.log("未找到Electronic Store条形码元素");
    }
}

window.createSimpleBarcode = createSimpleBarcode;

// 模板切换时调整高度
document.addEventListener('DOMContentLoaded', function () {
    const templateSelector = document.getElementById('template-selector');
    if (templateSelector) {
        templateSelector.addEventListener('change', function () {
            if (this.value === 'electronic-store-receipt-maker') {
                setTimeout(adjustElectronicStoreHeight, 100);
            }
        });
    }
});

// 提供全局函数供外部调用
window.adjustElectronicStoreHeight = adjustElectronicStoreHeight;

// 为 "General Grocery Store" 模板生成随机数
function generateGeneralGroceryRandomNumbers() {
    const templateId = '#general-grocery-store-template-for-food-grocery-meat-juices-and-bread-receipts';
    const template = document.querySelector(templateId);
    if (template) {
        template.querySelector('.purchase').innerText = generateRandomNumber(12);
        template.querySelector('.trace').innerText = generateRandomNumber(6);
        template.querySelector('.reference').innerText = generateRandomNumber(10);
        template.querySelector('.auth').innerText = generateRandomNumber(6);

        // 生成SROC编码
        const generateSrocPart = () => {
            const letter = String.fromCharCode(65 + Math.floor(Math.random() * 26)); // A-Z
            const numbers = generateRandomNumber(4);
            return `${letter}${numbers}`;
        };
        const srocCode = `${generateSrocPart()} ${generateSrocPart()} ${generateSrocPart()} ${generateSrocPart()}`;
        template.querySelector('.sroc').innerText = srocCode;
    }
}

// 为 "Online Furniture Shop" 模板生成随机数
function generateOnlineFurnitureShopRandomNumbers() {
    const templateId = '#online-receipts-for-furniture-shop';
    const template = document.querySelector(templateId);
    if (template) {
        // 生成SALE编号（20位数字）
        const saleElement = template.querySelector('.sale');
        if (saleElement) {
            saleElement.textContent = generateRandomNumber(20);
        }

        // 生成Auth No编号（6位数字）
        const authElement = template.querySelector('.authoriz');
        if (authElement) {
            authElement.textContent = generateRandomNumber(6);
        }

        // 生成AID编号（12位字符和数字混合）
        const aidElement = template.querySelector('.aid');
        if (aidElement) {
            let aidCode = '';
            for (let i = 0; i < 12; i++) {
                if (Math.random() < 0.5) {
                    // 生成数字
                    aidCode += Math.floor(Math.random() * 10);
                } else {
                    // 生成字母
                    aidCode += String.fromCharCode(65 + Math.floor(Math.random() * 26));
                }
            }
            aidElement.textContent = aidCode;
        }

        // 生成长数字编号（20位数字，类似85915860590807224707）
        const longNumberElement = template.querySelector('.long-number');
        if (longNumberElement) {
            longNumberElement.textContent = generateRandomNumber(20);
        }

        // 更新商品总数
        updateOnlineFurnitureShopItemCount();

        console.log('Online Furniture Shop模板随机数字已更新');
    }
}