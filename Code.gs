const SPREADSHEET_ID = "1eBhglF_U6sDz660MGjb7psh7IjZUQRZURgZjuXHd9lg";
const DRIVE_FOLDER_ID = "1x4yzn5uVeucbn1bQRmgiBee8_yvcFWYq";

const SHEET_PRODUCTS = "products";
const SHEET_STOCK = "stock";
const SHEET_ORDERS = "orders";
const SHEET_ORDER_ITEMS = "order_items";

const PRODUCTS_HEADERS = ["SKU", "name", "material", "color", "size", "price", "image", "product_id"];
const STOCK_HEADERS = ["SKU", "size", "quantity"];
const ORDERS_HEADERS = ["order_id", "date", "customer_name", "customer_phone", "customer_address", "status", "total_amount"];
const ORDER_ITEMS_HEADERS = ["order_id", "SKU", "size", "quantity"];

const ORDER_STATUS_PENDING = "pending_carrier";
const ORDER_STATUS_SENT = "sent_to_carrier";

function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || "{}");

    switch (data.action) {
      case "addProductFull":
        return json(addProductFull_(data));
      case "getProducts":
        return json(getProducts_(data));
      case "getOrders":
        return json(getOrders_());
      case "createOrder":
        return json(createOrder_(data));
      case "updateOrder":
        return json(updateOrder_(data));
      case "deleteOrder":
        return json(deleteOrder_(data));
      case "stockIn":
        return json(stockIn_(data));
      case "updateOrderStatus":
        return json(updateOrderStatus_(data));
      default:
        return json({ error: "Hành động không hợp lệ" });
    }
  } catch (err) {
    return json({ error: err.message || "Có lỗi không mong muốn xảy ra" });
  }
}

function doGet() {
  return json({
    status: "ok",
    actions: ["addProductFull", "getProducts", "getOrders", "createOrder", "updateOrder", "deleteOrder", "stockIn", "updateOrderStatus"]
  });
}

function json(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function addProductFull_(data) {
  const name = cleanString_(data.name);
  const material = cleanString_(data.material);
  const variants = normalizeVariants_(data.variants);
  const image = data.image || {};

  if (!name || !material) {
    throw new Error("Thiếu thông tin sản phẩm bắt buộc");
  }

  if (!variants.length) {
    throw new Error("Phải có ít nhất một biến thể màu và size");
  }

  if (!image.base64 || !image.fileName) {
    throw new Error("Bắt buộc phải có ảnh sản phẩm");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const resources = getResources_();
    const imageUrl = uploadImage_(image);
    const productId = generateProductId_();
    const productRows = createVariantRows_(productId, name, material, variants, imageUrl);
    const stockRows = productRows.map(function(row) {
      var variant = variants.find(function(item) {
        return item.color === row[3] && item.size === row[4];
      });
      return [row[0], row[4], variant ? variant.quantity : 0];
    });

    appendRows_(resources.productsSheet, productRows);
    appendRows_(resources.stockSheet, stockRows);

    invalidateProductsCache_();

    const allProductRows = getRows_(resources.productsSheet, PRODUCTS_HEADERS.length);
    const createdVariants = productRows.map(function(row, index) {
      return mapProductRow_(row, Number(stockRows[index][2] || 0));
    });

    return {
      status: "OK",
      product: mapCreatedProductPayload_(productId, name, material, imageUrl, createdVariants),
      variants: createdVariants,
      summary: {
        productCount: countProductGroups_(allProductRows),
        totalStock: sumStockSheet_(resources.stockSheet),
        orderCount: Math.max(0, resources.ordersSheet.getLastRow() - 1)
      }
    };
  } finally {
    lock.releaseLock();
  }
}

function getProducts_(options) {
  const includeOrders = !(options && options.includeOrders === false);
  const resources = getResources_();
  const productRows = getRows_(resources.productsSheet, PRODUCTS_HEADERS.length);
  const stockRows = getRows_(resources.stockSheet, STOCK_HEADERS.length);
  const stockMap = buildStockMap_(stockRows);
  const products = productRows
    .slice()
    .reverse()
    .map(function(row) {
      const key = stockKey_(row[0], row[4]);
      return mapProductRow_(row, stockMap[key] || 0);
    });
  const orderRows = includeOrders ? getRows_(resources.ordersSheet, ORDERS_HEADERS.length) : [];
  const orderItemRows = includeOrders ? getRows_(resources.orderItemsSheet, ORDER_ITEMS_HEADERS.length) : [];
  const productMap = includeOrders ? buildProductMap_(productRows) : {};
  const orders = includeOrders ? buildOrdersPayload_(orderRows, orderItemRows, productMap) : [];
  const orderCount = includeOrders ? orderRows.length : Math.max(0, resources.ordersSheet.getLastRow() - 1);

  const payload = {
    status: "OK",
    products: products,
    orders: orders,
    summary: {
      productCount: countProductGroups_(productRows),
      totalStock: stockRows.reduce(function(sum, row) { return sum + Number(row[2] || 0); }, 0),
      orderCount: orderCount
    }
  };

  return payload;
}

function getOrders_() {
  const resources = getResources_();
  const productRows = getRows_(resources.productsSheet, PRODUCTS_HEADERS.length);
  const orderRows = getRows_(resources.ordersSheet, ORDERS_HEADERS.length);
  const orderItemRows = getRows_(resources.orderItemsSheet, ORDER_ITEMS_HEADERS.length);
  const productMap = buildProductMap_(productRows);

  return {
    status: "OK",
    orders: buildOrdersPayload_(orderRows, orderItemRows, productMap),
    summary: {
      orderCount: orderRows.length
    }
  };
}

function stockIn_(data) {
  const productId = cleanString_(data.productId);
  const variants = normalizeVariants_(data.variants);
  const sku = cleanString_(data.sku);
  const size = cleanString_(data.size);
  const quantity = Number(data.quantity);

  if (productId && variants.length) {
    return stockInByProduct_(productId, variants);
  }

  if (!sku || !size) {
    throw new Error("Thiếu SKU hoặc size");
  }

  if (!Number.isInteger(quantity) || quantity < 1) {
    throw new Error("Số lượng phải lớn hơn hoặc bằng 1");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const resources = getResources_();
    const productRows = getRows_(resources.productsSheet, PRODUCTS_HEADERS.length);
    const productMap = buildProductMap_(productRows);
    const stockRows = getRows_(resources.stockSheet, STOCK_HEADERS.length);
    const index = findStockRowIndex_(stockRows, sku, size);
    let newQuantity = quantity;

    if (!productMap[sku]) {
      throw new Error("Không tìm thấy SKU: " + sku);
    }

    if (String(productMap[sku][4]) !== size) {
      throw new Error("Size không khớp với SKU: " + sku);
    }

    if (index >= 0) {
      newQuantity = Number(stockRows[index][2] || 0) + quantity;
      resources.stockSheet.getRange(index + 2, 3).setValue(newQuantity);
    } else {
      appendRow_(resources.stockSheet, [sku, size, quantity]);
    }

    invalidateProductsCache_();

    return {
      status: "OK",
      stock: {
        sku: sku,
        size: size,
        quantity: newQuantity
      },
      summary: {
        productCount: countProductGroups_(productRows),
        totalStock: sumStockSheet_(resources.stockSheet),
        orderCount: Math.max(0, resources.ordersSheet.getLastRow() - 1)
      }
    };
  } finally {
    lock.releaseLock();
  }
}

function stockInByProduct_(productId, variants) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const resources = getResources_();
    const productRows = getRows_(resources.productsSheet, PRODUCTS_HEADERS.length);
    const stockRows = getRows_(resources.stockSheet, STOCK_HEADERS.length);
    const groupRows = productRows.filter(function(row) {
      return getProductGroupKey_(row) === productId;
    });

    if (!groupRows.length) {
      throw new Error("Không tìm thấy mẫu áo dài " + productId);
    }

    const baseRow = groupRows[0];
    const variantIndexMap = buildProductVariantIndexByProductId_(productRows, productId);
    const stockIndex = buildStockIndex_(stockRows);
    const createdRows = [];
    const createdStocks = [];
    const updatedVariants = [];
    let nextVariantIndex = getNextVariantIndex_(groupRows);

    variants.forEach(function(variant) {
      const key = normalizeVariantKey_(variant.color, variant.size);
      const matchedRow = variantIndexMap[key];

      if (!Number.isInteger(variant.quantity) || variant.quantity < 1) {
        throw new Error("Mỗi dòng nhập kho phải có số lượng lớn hơn hoặc bằng 1");
      }

      if (matchedRow) {
        const sku = String(matchedRow[0]);
        const stockKey = stockKey_(sku, matchedRow[4]);
        const stockEntry = stockIndex[stockKey];
        const nextQuantity = (stockEntry ? stockEntry.quantity : 0) + variant.quantity;

        if (stockEntry) {
          resources.stockSheet.getRange(stockEntry.rowNumber, 3).setValue(nextQuantity);
          stockEntry.quantity = nextQuantity;
        } else {
          appendRow_(resources.stockSheet, [sku, matchedRow[4], variant.quantity]);
          stockIndex[stockKey] = {
            rowNumber: resources.stockSheet.getLastRow(),
            quantity: variant.quantity
          };
        }

        updatedVariants.push(mapProductRow_(matchedRow, nextQuantity));
        return;
      }

      const newRow = [
        buildVariantSku_(productId, nextVariantIndex),
        baseRow[1],
        baseRow[2],
        variant.color,
        variant.size,
        "",
        baseRow[6],
        productId
      ];
      const newStock = [newRow[0], newRow[4], variant.quantity];

      nextVariantIndex += 1;
      createdRows.push(newRow);
      createdStocks.push(newStock);
      updatedVariants.push(mapProductRow_(newRow, variant.quantity));
    });

    appendRows_(resources.productsSheet, createdRows);
    appendRows_(resources.stockSheet, createdStocks);
    invalidateProductsCache_();

    return {
      status: "OK",
      product: {
        productId: productId,
        name: baseRow[1],
        material: baseRow[2] || "",
        image: baseRow[6] || ""
      },
      variants: updatedVariants,
      summary: {
        productCount: countProductGroups_(getRows_(resources.productsSheet, PRODUCTS_HEADERS.length)),
        totalStock: sumStockSheet_(resources.stockSheet),
        orderCount: Math.max(0, resources.ordersSheet.getLastRow() - 1)
      }
    };
  } finally {
    lock.releaseLock();
  }
}

function createOrder_(data) {
  const customerName = cleanString_(data.customerName);
  const customerPhone = cleanPhone_(data.customerPhone);
  const customerAddress = cleanString_(data.customerAddress);
  const status = normalizeOrderStatus_(data.status);
  const totalAmount = Number(data.totalAmount);
  const items = normalizeOrderItems_(data.items);

  if (!customerName) {
    throw new Error("Thiếu tên khách hàng");
  }

  if (customerPhone.length < 8) {
    throw new Error("Số điện thoại khách hàng không hợp lệ");
  }

  if (!customerAddress) {
    throw new Error("Thiếu địa chỉ khách hàng");
  }

  if (!Number.isFinite(totalAmount) || totalAmount < 0) {
    throw new Error("Tổng tiền đơn hàng không hợp lệ");
  }

  if (!items.length) {
    throw new Error("Đơn hàng phải có ít nhất một dòng sản phẩm");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const resources = getResources_();
    const productRows = getRows_(resources.productsSheet, PRODUCTS_HEADERS.length);
    const stockRows = getRows_(resources.stockSheet, STOCK_HEADERS.length);
    const productMap = buildProductMap_(productRows);
    const stockIndex = buildStockIndex_(stockRows);

    items.forEach(function(item) {
      const product = productMap[item.sku];
      const stockEntry = stockIndex[stockKey_(item.sku, item.size)];

      if (!product) {
        throw new Error("Không tìm thấy SKU: " + item.sku);
      }

      if (String(product[4]) !== item.size) {
        throw new Error("Size không khớp với " + item.sku);
      }

      if (!stockEntry) {
        throw new Error("Không tìm thấy dòng tồn kho cho " + item.sku);
      }

      if (stockEntry.quantity < item.quantity) {
        throw new Error("Tồn kho không đủ cho " + item.sku);
      }
    });

    const orderId = generateOrderId_();
    const now = new Date().toISOString();
    const orderRow = [orderId, now, customerName, customerPhone, customerAddress, status, totalAmount];
    appendRow_(resources.ordersSheet, orderRow);

    const orderItemRows = items.map(function(item) {
      return [orderId, item.sku, item.size, item.quantity];
    });
    appendRows_(resources.orderItemsSheet, orderItemRows);

    const updatedStocks = [];
    items.forEach(function(item) {
      const entry = stockIndex[stockKey_(item.sku, item.size)];
      entry.quantity = entry.quantity - item.quantity;
      resources.stockSheet.getRange(entry.rowNumber, 3).setValue(entry.quantity);
      updatedStocks.push({
        sku: item.sku,
        size: item.size,
        quantity: entry.quantity
      });
    });

    invalidateProductsCache_();

    return {
      status: "OK",
      order: mapOrderRow_(orderRow, buildOrderItemsPayload_(orderItemRows, productMap)),
      updatedStocks: updatedStocks,
      summary: {
        productCount: countProductGroups_(productRows),
        totalStock: sumStockSheet_(resources.stockSheet),
        orderCount: Math.max(0, resources.ordersSheet.getLastRow() - 1)
      }
    };
  } finally {
    lock.releaseLock();
  }
}

function updateOrder_(data) {
  const orderId = cleanString_(data.orderId);
  const customerName = cleanString_(data.customerName);
  const customerPhone = cleanPhone_(data.customerPhone);
  const customerAddress = cleanString_(data.customerAddress);
  const status = normalizeOrderStatus_(data.status);
  const totalAmount = Number(data.totalAmount);
  const items = normalizeOrderItems_(data.items);

  if (!orderId) {
    throw new Error("Thiếu mã đơn hàng");
  }

  if (!customerName) {
    throw new Error("Thiếu tên khách hàng");
  }

  if (customerPhone.length < 8) {
    throw new Error("Số điện thoại khách hàng không hợp lệ");
  }

  if (!customerAddress) {
    throw new Error("Thiếu địa chỉ khách hàng");
  }

  if (!Number.isFinite(totalAmount) || totalAmount < 0) {
    throw new Error("Tổng tiền đơn hàng không hợp lệ");
  }

  if (!items.length) {
    throw new Error("Đơn hàng phải có ít nhất một dòng sản phẩm");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const resources = getResources_();
    const orderRows = getRows_(resources.ordersSheet, ORDERS_HEADERS.length);
    const orderIndex = findOrderRowIndex_(orderRows, orderId);

    if (orderIndex < 0) {
      throw new Error("Không tìm thấy đơn hàng " + orderId);
    }

    const productRows = getRows_(resources.productsSheet, PRODUCTS_HEADERS.length);
    const stockRows = getRows_(resources.stockSheet, STOCK_HEADERS.length);
    const orderItemRows = getRows_(resources.orderItemsSheet, ORDER_ITEMS_HEADERS.length);
    const productMap = buildProductMap_(productRows);
    const stockIndex = buildStockIndex_(stockRows);
    const existingItemEntries = findOrderItemEntries_(orderItemRows, orderId);
    const previousItems = existingItemEntries.map(function(entry) {
      return {
        sku: cleanString_(entry.row[1]),
        size: cleanString_(entry.row[2]),
        quantity: Number(entry.row[3] || 0)
      };
    });
    const changedStockKeys = {};

    previousItems.forEach(function(item) {
      const key = stockKey_(item.sku, item.size);
      const stockEntry = stockIndex[key];

      if (!stockEntry) {
        throw new Error("Không tìm thấy dòng tồn kho cho " + item.sku);
      }

      stockEntry.quantity += item.quantity;
      changedStockKeys[key] = true;
    });

    items.forEach(function(item) {
      const key = stockKey_(item.sku, item.size);
      const product = productMap[item.sku];
      const stockEntry = stockIndex[key];

      if (!product) {
        throw new Error("Không tìm thấy SKU: " + item.sku);
      }

      if (String(product[4]) !== item.size) {
        throw new Error("Size không khớp với " + item.sku);
      }

      if (!stockEntry) {
        throw new Error("Không tìm thấy dòng tồn kho cho " + item.sku);
      }

      if (stockEntry.quantity < item.quantity) {
        throw new Error("Tồn kho không đủ cho " + item.sku);
      }
    });

    items.forEach(function(item) {
      const key = stockKey_(item.sku, item.size);
      stockIndex[key].quantity -= item.quantity;
      changedStockKeys[key] = true;
    });

    persistStockChanges_(resources.stockSheet, stockIndex, Object.keys(changedStockKeys));

    resources.ordersSheet.getRange(orderIndex + 2, 3, 1, 5).setValues([[
      customerName,
      customerPhone,
      customerAddress,
      status,
      totalAmount
    ]]);

    deleteRowsByNumber_(resources.orderItemsSheet, existingItemEntries.map(function(entry) {
      return entry.rowNumber;
    }));

    const newOrderItemRows = items.map(function(item) {
      return [orderId, item.sku, item.size, item.quantity];
    });
    appendRows_(resources.orderItemsSheet, newOrderItemRows);

    invalidateProductsCache_();

    return {
      status: "OK",
      order: mapOrderRow_(
        [orderId, orderRows[orderIndex][1], customerName, customerPhone, customerAddress, status, totalAmount],
        buildOrderItemsPayload_(newOrderItemRows, productMap)
      ),
      updatedStocks: buildUpdatedStocksPayload_(stockIndex, Object.keys(changedStockKeys)),
      summary: {
        productCount: countProductGroups_(productRows),
        totalStock: sumStockSheet_(resources.stockSheet),
        orderCount: Math.max(0, resources.ordersSheet.getLastRow() - 1)
      }
    };
  } finally {
    lock.releaseLock();
  }
}

function deleteOrder_(data) {
  const orderId = cleanString_(data.orderId);

  if (!orderId) {
    throw new Error("Thiếu mã đơn hàng");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const resources = getResources_();
    const orderRows = getRows_(resources.ordersSheet, ORDERS_HEADERS.length);
    const orderIndex = findOrderRowIndex_(orderRows, orderId);

    if (orderIndex < 0) {
      throw new Error("Không tìm thấy đơn hàng " + orderId);
    }

    const productRows = getRows_(resources.productsSheet, PRODUCTS_HEADERS.length);
    const stockRows = getRows_(resources.stockSheet, STOCK_HEADERS.length);
    const orderItemRows = getRows_(resources.orderItemsSheet, ORDER_ITEMS_HEADERS.length);
    const stockIndex = buildStockIndex_(stockRows);
    const existingItemEntries = findOrderItemEntries_(orderItemRows, orderId);
    const changedStockKeys = {};

    existingItemEntries.forEach(function(entry) {
      const sku = cleanString_(entry.row[1]);
      const size = cleanString_(entry.row[2]);
      const quantity = Number(entry.row[3] || 0);
      const key = stockKey_(sku, size);
      const stockEntry = stockIndex[key];

      if (!stockEntry) {
        throw new Error("Không tìm thấy dòng tồn kho cho " + sku);
      }

      stockEntry.quantity += quantity;
      changedStockKeys[key] = true;
    });

    persistStockChanges_(resources.stockSheet, stockIndex, Object.keys(changedStockKeys));
    deleteRowsByNumber_(resources.orderItemsSheet, existingItemEntries.map(function(entry) {
      return entry.rowNumber;
    }));
    resources.ordersSheet.deleteRow(orderIndex + 2);

    invalidateProductsCache_();

    return {
      status: "OK",
      orderId: orderId,
      updatedStocks: buildUpdatedStocksPayload_(stockIndex, Object.keys(changedStockKeys)),
      summary: {
        productCount: countProductGroups_(productRows),
        totalStock: sumStockSheet_(resources.stockSheet),
        orderCount: Math.max(0, resources.ordersSheet.getLastRow() - 1)
      }
    };
  } finally {
    lock.releaseLock();
  }
}

function updateOrderStatus_(data) {
  const orderId = cleanString_(data.orderId);
  const status = normalizeOrderStatus_(data.status);

  if (!orderId) {
    throw new Error("Thiếu mã đơn hàng");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const resources = getResources_();
    const orderRows = getRows_(resources.ordersSheet, ORDERS_HEADERS.length);
    const orderIndex = findOrderRowIndex_(orderRows, orderId);

    if (orderIndex < 0) {
      throw new Error("Không tìm thấy đơn hàng " + orderId);
    }

    resources.ordersSheet.getRange(orderIndex + 2, 6).setValue(status);
    orderRows[orderIndex][5] = status;

    const productRows = getRows_(resources.productsSheet, PRODUCTS_HEADERS.length);
    const orderItemRows = getRows_(resources.orderItemsSheet, ORDER_ITEMS_HEADERS.length);
    const productMap = buildProductMap_(productRows);
    const orderItems = buildOrderItemsPayload_(
      orderItemRows.filter(function(row) { return String(row[0]) === orderId; }),
      productMap
    );

    invalidateProductsCache_();

    return {
      status: "OK",
      order: mapOrderRow_(orderRows[orderIndex], orderItems),
      summary: {
        productCount: countProductGroups_(productRows),
        totalStock: sumStockSheet_(resources.stockSheet),
        orderCount: Math.max(0, resources.ordersSheet.getLastRow() - 1)
      }
    };
  } finally {
    lock.releaseLock();
  }
}

function getResources_() {
  if (SPREADSHEET_ID === "YOUR_SPREADSHEET_ID" || DRIVE_FOLDER_ID === "YOUR_DRIVE_FOLDER_ID") {
    throw new Error("Hãy cấu hình SPREADSHEET_ID và DRIVE_FOLDER_ID trước");
  }

  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  return {
    spreadsheet: spreadsheet,
    productsSheet: getOrCreateSheet_(spreadsheet, SHEET_PRODUCTS, PRODUCTS_HEADERS),
    stockSheet: getOrCreateSheet_(spreadsheet, SHEET_STOCK, STOCK_HEADERS),
    ordersSheet: getOrCreateSheet_(spreadsheet, SHEET_ORDERS, ORDERS_HEADERS),
    orderItemsSheet: getOrCreateSheet_(spreadsheet, SHEET_ORDER_ITEMS, ORDER_ITEMS_HEADERS)
  };
}

function getOrCreateSheet_(spreadsheet, name, headers) {
  let sheet = spreadsheet.getSheetByName(name);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
  }

  if (name === SHEET_PRODUCTS) {
    migrateProductsSheet_(sheet);
  }

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  const currentHeaders = headerRange.getValues()[0];
  const hasDifferentHeaders = headers.some(function(header, index) {
    return String(currentHeaders[index] || "") !== header;
  });

  if (sheet.getLastRow() === 0 || hasDifferentHeaders) {
    headerRange.setValues([headers]);
  }

  if (name === SHEET_PRODUCTS) {
    repairMisalignedProductRows_(sheet);
    populateLegacyProductIds_(sheet);
  }

  return sheet;
}

function migrateProductsSheet_(sheet) {
  if (sheet.getLastRow() === 0) {
    return;
  }

  var currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 8)).getValues()[0];
  var hasMaterialHeader = String(currentHeaders[2] || "") === "material";
  var isLegacyProducts =
    String(currentHeaders[0] || "") === "SKU" &&
    String(currentHeaders[1] || "") === "name" &&
    String(currentHeaders[2] || "") === "color" &&
    String(currentHeaders[3] || "") === "size" &&
    String(currentHeaders[4] || "") === "price" &&
    String(currentHeaders[5] || "") === "image";

  if (hasMaterialHeader) {
    populateLegacyProductIds_(sheet);
    return;
  }

  if (!isLegacyProducts) {
    return;
  }

  sheet.insertColumnAfter(2);
  sheet.getRange(1, 1, 1, PRODUCTS_HEADERS.length).setValues([PRODUCTS_HEADERS]);
  populateLegacyProductIds_(sheet);
}

function populateLegacyProductIds_(sheet) {
  if (sheet.getLastRow() <= 1) {
    return;
  }

  var productIdRange = sheet.getRange(2, 8, sheet.getLastRow() - 1, 1);
  var productIdValues = productIdRange.getValues();
  var skuValues = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  var changed = false;

  for (var i = 0; i < productIdValues.length; i += 1) {
    if (!cleanString_(productIdValues[i][0])) {
      productIdValues[i][0] = cleanString_(skuValues[i][0]);
      changed = true;
    }
  }

  if (changed) {
    productIdRange.setValues(productIdValues);
  }
}

function repairMisalignedProductRows_(sheet) {
  if (sheet.getLastRow() <= 1) {
    return;
  }

  var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, PRODUCTS_HEADERS.length);
  var values = range.getValues();
  var changed = false;

  values = values.map(function(row) {
    var priceColumnLooksLikeImage = looksLikeImageValue_(row[5]);
    var legacyPrice = Number(row[4]);

    if (priceColumnLooksLikeImage && !cleanString_(row[6]) && Number.isFinite(legacyPrice)) {
      changed = true;
      return [
        row[0],
        row[1],
        "",
        row[2],
        row[3],
        legacyPrice,
        row[5],
        cleanString_(row[7]) || cleanString_(row[0])
      ];
    }

    return row;
  });

  if (changed) {
    range.setValues(values);
  }
}

function getRows_(sheet, width) {
  if (sheet.getLastRow() <= 1) {
    return [];
  }

  return sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getValues();
}

function appendRow_(sheet, row) {
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
}

function appendRows_(sheet, rows) {
  if (!rows.length) {
    return;
  }

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function sumStockSheet_(sheet) {
  return getRows_(sheet, STOCK_HEADERS.length).reduce(function(sum, row) {
    return sum + Number(row[2] || 0);
  }, 0);
}

function buildStockMap_(rows) {
  return rows.reduce(function(map, row) {
    map[stockKey_(row[0], row[1])] = Number(row[2] || 0);
    return map;
  }, {});
}

function buildStockIndex_(rows) {
  return rows.reduce(function(map, row, index) {
    map[stockKey_(row[0], row[1])] = {
      rowNumber: index + 2,
      quantity: Number(row[2] || 0)
    };
    return map;
  }, {});
}

function buildProductMap_(rows) {
  return rows.reduce(function(map, row) {
    map[row[0]] = row;
    return map;
  }, {});
}

function buildProductVariantIndexByProductId_(rows, productId) {
  return rows.reduce(function(map, row, index) {
    if (getProductGroupKey_(row) !== productId) {
      return map;
    }

    row.__rowNumber = index + 2;
    map[normalizeVariantKey_(row[3], row[4])] = row;
    return map;
  }, {});
}

function countProductGroups_(rows) {
  var groups = rows.reduce(function(map, row) {
    map[getProductGroupKey_(row)] = true;
    return map;
  }, {});

  return Object.keys(groups).length;
}

function buildOrdersPayload_(orderRows, orderItemRows, productMap) {
  const itemsByOrder = buildOrderItemsByOrder_(orderItemRows, productMap);
  return orderRows
    .slice()
    .reverse()
    .map(function(row) {
      return mapOrderRow_(row, itemsByOrder[row[0]] || []);
    });
}

function buildOrderItemsByOrder_(orderItemRows, productMap) {
  return orderItemRows.reduce(function(map, row) {
    const orderId = String(row[0] || "");
    map[orderId] = map[orderId] || [];
    map[orderId].push(mapOrderItemPayload_(row, productMap[row[1]]));
    return map;
  }, {});
}

function buildOrderItemsPayload_(rows, productMap) {
  return rows.map(function(row) {
    return mapOrderItemPayload_(row, productMap[row[1]]);
  });
}

function findStockRowIndex_(rows, sku, size) {
  for (var i = 0; i < rows.length; i += 1) {
    if (String(rows[i][0]) === sku && String(rows[i][1]) === size) {
      return i;
    }
  }

  return -1;
}

function findOrderRowIndex_(rows, orderId) {
  for (var i = 0; i < rows.length; i += 1) {
    if (String(rows[i][0]) === orderId) {
      return i;
    }
  }

  return -1;
}

function findOrderItemEntries_(rows, orderId) {
  return rows.reduce(function(items, row, index) {
    if (String(row[0]) === orderId) {
      items.push({
        row: row,
        rowNumber: index + 2
      });
    }

    return items;
  }, []);
}

function persistStockChanges_(sheet, stockIndex, keys) {
  keys.forEach(function(key) {
    if (stockIndex[key]) {
      sheet.getRange(stockIndex[key].rowNumber, 3).setValue(stockIndex[key].quantity);
    }
  });
}

function buildUpdatedStocksPayload_(stockIndex, keys) {
  return keys.map(function(key) {
    const parts = String(key).split("|");
    return {
      sku: parts[0] || "",
      size: parts[1] || "",
      quantity: stockIndex[key] ? Number(stockIndex[key].quantity || 0) : 0
    };
  });
}

function deleteRowsByNumber_(sheet, rowNumbers) {
  rowNumbers
    .slice()
    .sort(function(left, right) { return right - left; })
    .forEach(function(rowNumber) {
      sheet.deleteRow(rowNumber);
    });
}

function uploadImage_(image) {
  const mimeType = cleanString_(image.mimeType) || "image/jpeg";
  const bytes = Utilities.base64Decode(image.base64);

  if (["image/jpeg", "image/png", "image/webp"].indexOf(mimeType) === -1) {
    throw new Error("Định dạng ảnh không được hỗ trợ");
  }

  if (bytes.length > 400 * 1024) {
    throw new Error("Ảnh sau khi nén vẫn còn quá lớn");
  }

  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const blob = Utilities.newBlob(bytes, mimeType, sanitizeFileName_(image.fileName));
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return "https://lh3.googleusercontent.com/d/" + file.getId();
}

function normalizeOrderItems_(items) {
  if (!Array.isArray(items)) {
    return [];
  }

  const map = {};
  items.forEach(function(item) {
    const sku = cleanString_(item.sku);
    const size = cleanString_(item.size);
    const quantity = Number(item.quantity);

    if (!sku || !size) {
      throw new Error("Mỗi dòng đơn hàng phải có SKU và size");
    }

    if (!Number.isInteger(quantity) || quantity < 1) {
      throw new Error("Số lượng của mỗi dòng đơn hàng phải lớn hơn hoặc bằng 1");
    }

    const key = stockKey_(sku, size);
    map[key] = map[key] || { sku: sku, size: size, quantity: 0 };
    map[key].quantity += quantity;
  });

  return Object.keys(map).map(function(key) {
    return map[key];
  });
}

function mapProductRow_(row, currentStock) {
  return {
    sku: row[0],
    name: row[1],
    material: row[2] || "",
    color: row[3],
    size: row[4],
    price: 0,
    image: row[6],
    productId: row[7] || row[0],
    currentStock: Number(currentStock || 0)
  };
}

function mapCreatedProductPayload_(productId, name, material, imageUrl, variants) {
  return {
    productId: productId,
    name: name,
    material: material,
    image: imageUrl,
    colors: variants.map(function(item) { return item.color; }).filter(uniqueValue_),
    sizes: variants.map(function(item) { return item.size; }).filter(uniqueValue_),
    variantCount: variants.length
  };
}

function mapOrderRow_(row, items) {
  return {
    orderId: row[0],
    date: row[1],
    customerName: row[2] || "",
    customerPhone: row[3] || "",
    customerAddress: row[4] || "",
    status: normalizeOrderStatus_(row[5]),
    totalAmount: Number(row[6] || 0),
    items: items || []
  };
}

function mapOrderItemPayload_(row, productRow) {
  return {
    sku: row[1],
    size: row[2],
    quantity: Number(row[3] || 0),
    name: productRow ? productRow[1] : "",
    material: productRow ? productRow[2] : "",
    color: productRow ? productRow[3] : "",
    image: productRow ? productRow[6] : ""
  };
}

function stockKey_(sku, size) {
  return String(sku) + "|" + String(size);
}

function getProductGroupKey_(row) {
  return cleanString_(row[7]) || cleanString_(row[0]);
}

function createVariantRows_(productId, name, material, variants, imageUrl) {
  return variants.map(function(variant, index) {
    return [
      buildVariantSku_(productId, index + 1),
      name,
      material,
      variant.color,
      variant.size,
      "",
      imageUrl,
      productId
    ];
  });
}

function generateProductId_() {
  return "MAU-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyMMdd-HHmmss") + "-" + Math.floor(Math.random() * 900 + 100);
}

function buildVariantSku_(productId, index) {
  return productId + "-V" + padNumber_(index, 2);
}

function getNextVariantIndex_(rows) {
  var maxIndex = 0;

  rows.forEach(function(row) {
    var match = String(row[0] || "").match(/-V(\d+)$/);
    if (match) {
      maxIndex = Math.max(maxIndex, Number(match[1] || 0));
    }
  });

  return maxIndex + 1;
}

function generateOrderId_() {
  return "ORD-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyMMdd-HHmmss") + "-" + Math.floor(Math.random() * 900 + 100);
}

function padNumber_(value, width) {
  var output = String(value);

  while (output.length < width) {
    output = "0" + output;
  }

  return output;
}

function sanitizeFileName_(name) {
  return cleanString_(name).replace(/[^\w.\-]/g, "_") || "image.jpg";
}

function cleanString_(value) {
  return String(value || "").trim();
}

function looksLikeImageValue_(value) {
  return /^https?:\/\//i.test(cleanString_(value)) || /^data:image\//i.test(cleanString_(value));
}

function normalizeVariantKey_(color, size) {
  return (cleanString_(color) + "|" + cleanString_(size)).toLowerCase();
}

function normalizeVariants_(variants) {
  if (!Array.isArray(variants)) {
    return [];
  }

  var seen = {};

  return variants.reduce(function(items, variant) {
    var color = cleanString_(variant && variant.color);
    var size = cleanString_(variant && variant.size);
    var quantity = Number(variant && variant.quantity);
    var key = (color + "|" + size).toLowerCase();

    if (!color || !size) {
      throw new Error("Mỗi biến thể phải có màu và size");
    }

    if (!Number.isInteger(quantity) || quantity < 0) {
      quantity = 0;
    }

    if (!seen[key]) {
      seen[key] = true;
      items.push({
        color: color,
        size: size,
        quantity: quantity
      });
    }

    return items;
  }, []);
}

function cleanPhone_(value) {
  return String(value || "").replace(/[^\d+]/g, "");
}

function uniqueValue_(value, index, array) {
  return array.indexOf(value) === index;
}

function normalizeOrderStatus_(value) {
  return value === ORDER_STATUS_SENT ? ORDER_STATUS_SENT : ORDER_STATUS_PENDING;
}

function invalidateProductsCache_() {
  // App reads directly from Sheets now, so there is no backend cache to clear.
}
