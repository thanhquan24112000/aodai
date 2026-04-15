const CONFIG = {
  API_URL: "https://script.google.com/macros/s/AKfycbyXruhnDCVCY9Xzp2zwbcmQV6W5SNwPKxPjHptbVtL9NgVWzsdBm6chCXKW3grQMwzQiQ/exec",
  MAX_IMAGE_WIDTH: 500,
  IMAGE_QUALITY: 0.7,
  TARGET_IMAGE_BYTES: 300 * 1024,
  ORDERS_PAGE_SIZE: 10
};

const ORDER_STATUS = {
  PENDING: "pending_carrier",
  SENT: "sent_to_carrier"
};

const SECTION_META = {
  dashboard: {
    title: "Trang chủ",
    description: "Xem nhanh số mẫu, tồn kho và đơn hàng, rồi chạm vào đúng khu vực cần xử lý."
  },
  products: {
    title: "Danh sách sản phẩm",
    description: ""
  },
  inventory: {
    title: "Kho hàng",
    description: "Chọn đúng biến thể màu và size để nhập kho và kiểm tra tồn hiện tại."
  },
  orders: {
    title: "Danh sách đơn hàng",
    description: ""
  }
};

const money = new Intl.NumberFormat("vi-VN");
const state = {
  products: [],
  summary: {
    productCount: 0,
    totalStock: 0,
    orderCount: 0
  },
  orders: [],
  productVariants: [],
  stockVariants: [],
  orderDraft: [],
  stockTargetProductId: "",
  stockPreferredSku: "",
  orderModalMode: "create",
  editingOrderId: "",
  orderOriginalItems: [],
  pendingImage: null,
  lastSync: "",
  isBusy: false,
  isLoadingOrders: false,
  hasLoaded: false,
  ordersLoaded: false,
  orderFilter: "pending",
  visibleOrdersCount: CONFIG.ORDERS_PAGE_SIZE,
  selectedOrderSku: "",
  filter: "",
  activeSection: "dashboard",
  activeModal: ""
};

const el = {};

document.addEventListener("DOMContentLoaded", init);

async function init() {
  cacheElements();
  bindEvents();
  render();
  disableClientCache();
  await loadProducts({ includeOrders: true });
}

function cacheElements() {
  el.sectionNavs = Array.from(document.querySelectorAll("[data-section-nav]"));
  el.navButtons = Array.from(document.querySelectorAll("[data-section-target]"));
  el.sectionPanels = Array.from(document.querySelectorAll("[data-section-panel]"));
  el.quickButtons = Array.from(document.querySelectorAll("[data-go-section]"));
  el.helpToggleButtons = Array.from(document.querySelectorAll("[data-help-toggle]"));
  el.helpPanels = Array.from(document.querySelectorAll("[data-help-panel]"));
  el.sectionTitle = document.getElementById("sectionTitle");
  el.sectionDescription = document.getElementById("sectionDescription");
  el.toastViewport = document.getElementById("toastViewport");

  el.refreshButton = document.getElementById("refreshButton");
  el.syncNote = document.getElementById("syncNote");
  el.totalProducts = document.getElementById("totalProducts");
  el.totalStock = document.getElementById("totalStock");
  el.totalOrders = document.getElementById("totalOrders");
  el.catalogSearch = document.getElementById("catalogSearch");
  el.openProductModalButton = document.getElementById("openProductModalButton");
  el.productModal = document.getElementById("productModal");
  el.closeProductModalButton = document.getElementById("closeProductModalButton");
  el.productDetailModal = document.getElementById("productDetailModal");
  el.closeProductDetailButton = document.getElementById("closeProductDetailButton");
  el.productDetailTitle = document.getElementById("productDetailTitle");
  el.productDetailBody = document.getElementById("productDetailBody");
  el.stockModal = document.getElementById("stockModal");
  el.closeStockModalButton = document.getElementById("closeStockModalButton");
  el.orderModal = document.getElementById("orderModal");
  el.openOrderModalButton = document.getElementById("openOrderModalButton");
  el.closeOrderModalButton = document.getElementById("closeOrderModalButton");
  el.orderModalKicker = document.getElementById("orderModalKicker");
  el.orderModalTitle = document.getElementById("orderModalTitle");
  el.orderModalMeta = document.getElementById("orderModalMeta");
  el.orderDetailModal = document.getElementById("orderDetailModal");
  el.closeOrderDetailButton = document.getElementById("closeOrderDetailButton");
  el.orderDetailTitle = document.getElementById("orderDetailTitle");
  el.orderDetailBody = document.getElementById("orderDetailBody");

  el.dashboardRecentList = document.getElementById("dashboardRecentList");
  el.lowStockList = document.getElementById("lowStockList");
  el.inventoryList = document.getElementById("inventoryList");
  el.ordersList = document.getElementById("ordersList");
  el.ordersStatusTabs = document.getElementById("ordersStatusTabs");

  el.productForm = document.getElementById("productForm");
  el.productName = document.getElementById("productName");
  el.productMaterial = document.getElementById("productMaterial");
  el.variantPreview = document.getElementById("variantPreview");
  el.variantColor = document.getElementById("variantColor");
  el.variantSize = document.getElementById("variantSize");
  el.variantQuantity = document.getElementById("variantQuantity");
  el.addVariantButton = document.getElementById("addVariantButton");
  el.variantDraftList = document.getElementById("variantDraftList");
  el.productImage = document.getElementById("productImage");
  el.imagePreview = document.getElementById("imagePreview");
  el.productSubmitButton = document.getElementById("productSubmitButton");
  el.productList = document.getElementById("productList");
  el.productDraftCount = document.getElementById("productDraftCount");
  el.productDraftStock = document.getElementById("productDraftStock");

  el.stockForm = document.getElementById("stockForm");
  el.stockModalProductCard = document.getElementById("stockModalProductCard");
  el.stockQuantity = document.getElementById("stockQuantity");
  el.stockColor = document.getElementById("stockColor");
  el.stockSize = document.getElementById("stockSize");
  el.stockVariantHint = document.getElementById("stockVariantHint");
  el.addStockVariantButton = document.getElementById("addStockVariantButton");
  el.stockVariantPreview = document.getElementById("stockVariantPreview");
  el.stockVariantDraftList = document.getElementById("stockVariantDraftList");
  el.stockDraftCount = document.getElementById("stockDraftCount");
  el.stockDraftQuantity = document.getElementById("stockDraftQuantity");
  el.stockSubmitButton = document.getElementById("stockSubmitButton");
  el.stockCurrentBadge = document.getElementById("stockCurrentBadge");

  el.orderItemForm = document.getElementById("orderItemForm");
  el.orderForm = document.getElementById("orderForm");
  el.customerName = document.getElementById("customerName");
  el.customerPhone = document.getElementById("customerPhone");
  el.customerAddress = document.getElementById("customerAddress");
  el.orderStatus = document.getElementById("orderStatus");
  el.orderTotalAmount = document.getElementById("orderTotalAmount");
  el.orderProductSearch = document.getElementById("orderProductSearch");
  el.orderProductResults = document.getElementById("orderProductResults");
  el.orderQuantity = document.getElementById("orderQuantity");
  el.addItemButton = document.getElementById("addItemButton");
  el.orderHint = document.getElementById("orderHint");
  el.orderDraft = document.getElementById("orderDraft");
  el.orderDraftTotal = document.getElementById("orderDraftTotal");
  el.orderDraftSummary = document.getElementById("orderDraftSummary");
  el.submitOrderButton = document.getElementById("submitOrderButton");

  el.loadingOverlay = document.getElementById("loadingOverlay");
  el.loadingTitle = document.getElementById("loadingTitle");
  el.loadingMessage = document.getElementById("loadingMessage");
}

function bindEvents() {
  el.sectionNavs.forEach((nav) => {
    nav.addEventListener("click", handleSectionNavigation);
  });
  el.quickButtons.forEach((button) => {
    button.addEventListener("click", () => switchSection(button.dataset.goSection));
  });
  el.helpToggleButtons.forEach((button) => {
    button.addEventListener("click", handleHelpToggle);
  });
  if (el.toastViewport) {
    el.toastViewport.addEventListener("click", handleToastActions);
  }

  el.refreshButton.addEventListener("click", handleRefresh);
  el.catalogSearch.addEventListener("input", handleCatalogSearch);
  el.openProductModalButton.addEventListener("click", openProductModal);
  el.closeProductModalButton.addEventListener("click", closeActiveModal);
  el.closeProductDetailButton.addEventListener("click", closeActiveModal);
  el.closeStockModalButton.addEventListener("click", closeActiveModal);
  el.openOrderModalButton.addEventListener("click", () => openOrderModal({ mode: "create" }));
  el.closeOrderModalButton.addEventListener("click", closeActiveModal);
  el.closeOrderDetailButton.addEventListener("click", closeActiveModal);
  el.productModal.addEventListener("click", handleModalBackdropClick);
  el.productDetailModal.addEventListener("click", handleModalBackdropClick);
  el.stockModal.addEventListener("click", handleModalBackdropClick);
  el.orderModal.addEventListener("click", handleModalBackdropClick);
  el.orderDetailModal.addEventListener("click", handleModalBackdropClick);
  el.productImage.addEventListener("change", handleImageChange);
  el.addVariantButton.addEventListener("click", handleAddVariant);
  el.variantDraftList.addEventListener("click", handleVariantDraftActions);
  el.productForm.addEventListener("submit", handleAddProduct);
  el.stockForm.addEventListener("submit", handleStockIn);
  el.stockColor.addEventListener("input", syncStockVariantDefaults);
  el.stockSize.addEventListener("input", syncStockVariantDefaults);
  el.addStockVariantButton.addEventListener("click", handleAddStockVariant);
  el.stockVariantDraftList.addEventListener("click", handleStockVariantDraftActions);
  el.orderItemForm.addEventListener("submit", handleAddOrderItem);
  el.orderProductSearch.addEventListener("input", handleOrderProductSearch);
  el.orderProductResults.addEventListener("click", handleOrderProductResultsClick);
  el.orderDraft.addEventListener("click", handleDraftActions);
  el.productList.addEventListener("click", handleProductCardClick);
  el.productDetailBody.addEventListener("click", handleProductDetailClick);
  el.inventoryList.addEventListener("click", handleInventoryListClick);
  el.ordersList.addEventListener("click", handleOrdersListClick);
  el.ordersStatusTabs.addEventListener("click", handleOrderFilterChange);
  el.orderDetailBody.addEventListener("click", handleOrdersListClick);
  el.lowStockList.addEventListener("click", handleLowStockListClick);
  el.submitOrderButton.addEventListener("click", handleSubmitOrder);
  window.addEventListener("keydown", handleEscapeKey);
  window.addEventListener("online", () => loadProducts({ silent: true }));
}

async function handleRefresh() {
  await loadProducts({
    includeOrders: state.activeSection === "orders" || state.activeSection === "dashboard"
  });
}

async function loadProducts(options = {}) {
  const silent = Boolean(options.silent);
  const includeOrders = Boolean(options.includeOrders);
  const showLoading = !silent;

  if (!silent && !state.products.length) {
    renderSkeletons();
  }

  if (showLoading) {
    setLoadingState(true, includeOrders ? "sync-all" : "sync-products");
  }

  try {
    const data = await postAction("getProducts", {
      includeOrders
    });
    const productPayload = Array.isArray(data) ? data : (data.products || data.data || []);
    const orderPayload = Array.isArray(data) ? [] : (data.orders || []);
    state.products = productPayload.map(normalizeProduct).filter(isRenderableProduct);
    if (includeOrders) {
      state.orders = orderPayload.map(normalizeOrder);
      state.ordersLoaded = true;
      resetVisibleOrdersCount();
    }
    state.summary = normalizeSummary(data.summary || {});
    state.lastSync = new Date().toISOString();
    state.hasLoaded = true;
    render();

    if (!silent) {
      setStatus("Đồng bộ danh mục thành công.", "success");
    }
  } catch (error) {
    render();
    if (!state.products.length) {
      setStatus(error.message || "Không thể tải danh mục sản phẩm.", "error");
    } else if (!silent) {
      setStatus(error.message || "Không thể tải dữ liệu mới. App đang giữ dữ liệu hiện tại trên màn hình.", "error");
    }
  } finally {
    if (showLoading) {
      setLoadingState(false, includeOrders ? "sync-all" : "sync-products");
    }
  }
}

async function loadOrders(options = {}) {
  const silent = Boolean(options.silent);

  if (state.isLoadingOrders) {
    return;
  }

  state.isLoadingOrders = true;

  if (!silent && !state.orders.length) {
    renderOrdersLoading();
  }

  if (!silent) {
    setLoadingState(true, "sync-orders");
  }

  try {
    const data = await postAction("getOrders");
    state.orders = (data.orders || []).map(normalizeOrder);
    resetVisibleOrdersCount();
    state.summary = normalizeSummary({
      ...state.summary,
      ...(data.summary || {})
    });
    state.ordersLoaded = true;
    state.lastSync = new Date().toISOString();
    state.isLoadingOrders = false;
    render();

    if (!silent) {
      setStatus("Đã tải danh sách đơn hàng.", "success");
    }
  } catch (error) {
    state.isLoadingOrders = false;
    render();
    if (!silent) {
      setStatus(error.message || "Không thể tải danh sách đơn hàng.", "error");
    }
  } finally {
    if (!silent) {
      setLoadingState(false, "sync-orders");
    }
  }
}

function handleSectionNavigation(event) {
  const button = event.target.closest("[data-section-target]");
  if (!button) {
    return;
  }

  switchSection(button.dataset.sectionTarget);
}

function handleOrderFilterChange(event) {
  const button = event.target.closest("[data-order-filter]");
  if (!button) {
    return;
  }

  state.orderFilter = button.dataset.orderFilter === "sent" ? "sent" : "pending";
  resetVisibleOrdersCount();
  renderOrdersTabs();
  renderOrdersList();
}

function resetVisibleOrdersCount() {
  state.visibleOrdersCount = CONFIG.ORDERS_PAGE_SIZE;
}

function switchSection(sectionKey) {
  if (sectionKey === "inventory") {
    sectionKey = "products";
  }

  if (!SECTION_META[sectionKey]) {
    return;
  }

  state.activeSection = sectionKey;
  closeHelpPanels();
  renderSectionState();

  if (sectionKey === "orders" && !state.ordersLoaded) {
    loadOrders();
  }
}

function handleHelpToggle(event) {
  const sectionKey = event.currentTarget.dataset.helpToggle;
  if (!sectionKey) {
    return;
  }

  const button = getHelpToggleButton(sectionKey);
  const panel = getHelpPanel(sectionKey);
  if (!button || !panel) {
    return;
  }

  const isOpen = button.getAttribute("aria-expanded") === "true";
  setHelpPanelState(sectionKey, !isOpen);
}

function getHelpToggleButton(sectionKey) {
  return el.helpToggleButtons.find((button) => button.dataset.helpToggle === sectionKey) || null;
}

function getHelpPanel(sectionKey) {
  return el.helpPanels.find((panel) => panel.dataset.helpPanel === sectionKey) || null;
}

function setHelpPanelState(sectionKey, isOpen) {
  const button = getHelpToggleButton(sectionKey);
  const panel = getHelpPanel(sectionKey);
  if (!button || !panel) {
    return;
  }

  panel.hidden = !isOpen;
  button.classList.toggle("is-open", isOpen);
  button.setAttribute("aria-expanded", String(isOpen));
  button.textContent = isOpen ? "Ẩn hỗ trợ" : "Hỗ trợ";
}

function closeHelpPanels() {
  el.helpPanels.forEach((panel) => {
    panel.hidden = true;
  });

  el.helpToggleButtons.forEach((button) => {
    button.classList.remove("is-open");
    button.setAttribute("aria-expanded", "false");
    button.textContent = "Hỗ trợ";
  });
}

function handleToastActions(event) {
  const closeButton = event.target.closest("[data-toast-close]");
  if (!closeButton) {
    return;
  }

  dismissToast(closeButton.closest(".toast"));
}

function handleCatalogSearch(event) {
  state.filter = event.target.value.trim().toLowerCase();
  renderProducts();
}

function handleModalBackdropClick(event) {
  if (event.target === event.currentTarget) {
    closeActiveModal();
  }
}

function handleEscapeKey(event) {
  if (event.key === "Escape" && state.activeModal) {
    closeActiveModal();
  }
}

function openProductModal() {
  if (state.activeModal && state.activeModal !== "product") {
    closeActiveModal();
  }

  state.activeModal = "product";
  el.productModal.hidden = false;
  el.productModal.setAttribute("aria-hidden", "false");
  document.body.classList.add("modal-open");
  syncVariantPriceDefault();
  renderVariantPreview();
  renderVariantDraft();
  requestAnimationFrame(() => {
    el.productName.focus();
  });
}

function openProductDetailModal(productId) {
  const product = getCatalogProducts().find((entry) => entry.productId === productId);
  if (!product) {
    setStatus("Không tìm thấy sản phẩm.", "error");
    return;
  }

  if (state.activeModal && state.activeModal !== "product-detail") {
    closeActiveModal();
  }

  state.activeModal = "product-detail";
  el.productDetailModal.hidden = false;
  el.productDetailModal.setAttribute("aria-hidden", "false");
  document.body.classList.add("modal-open");
  renderProductDetail(product);
}

function openStockModal(productId, preferredSku = "") {
  const catalogProduct = getCatalogProducts().find((product) => product.productId === productId);
  if (!catalogProduct) {
    return;
  }

  if (state.activeModal && state.activeModal !== "stock") {
    closeActiveModal();
  }

  state.stockTargetProductId = productId;
  state.stockPreferredSku = preferredSku;
  state.stockVariants = [];
  state.activeModal = "stock";
  el.stockModal.hidden = false;
  el.stockModal.setAttribute("aria-hidden", "false");
  document.body.classList.add("modal-open");
  renderStockModalProductCard(preferredSku);
  resetStockComposer({ preferredVariant: preferredSku ? getSelectedProduct(preferredSku) : null });
  renderStockBadge();
  renderStockVariantPreview();
  renderStockVariantDraft();
  requestAnimationFrame(() => {
    el.stockColor.focus();
  });
}

function openOrderModal(options = {}) {
  const mode = options.mode === "edit" ? "edit" : "create";
  const orderId = options.orderId || "";
  const preferredSku = options.preferredSku || "";
  const order = mode === "edit"
    ? state.orders.find((entry) => entry.orderId === orderId)
    : null;

  if (mode === "edit" && !order) {
    setStatus("Không tìm thấy đơn hàng để chỉnh sửa.", "error");
    return;
  }

  if (state.activeModal && state.activeModal !== "order") {
    closeActiveModal();
  }

  state.orderModalMode = mode;
  state.editingOrderId = order ? order.orderId : "";
  state.orderOriginalItems = order ? order.items.map(cloneOrderItem) : [];
  state.orderDraft = order ? order.items.map(cloneOrderItem) : [];
  state.activeModal = "order";

  el.orderModal.hidden = false;
  el.orderModal.setAttribute("aria-hidden", "false");
  document.body.classList.add("modal-open");

  el.orderModalKicker.textContent = mode === "edit" ? "Chỉnh sửa đơn hàng" : "Đơn hàng";
  el.orderModalTitle.textContent = mode === "edit" ? `Sửa đơn ${order.orderId}` : "Tạo đơn hàng";
  el.orderModalMeta.textContent = mode === "edit"
    ? "Bạn có thể sửa khách hàng, trạng thái, các dòng áo và tổng tiền của đơn này."
    : "Điền thông tin khách hàng, thêm các áo vào đơn và nhập tổng tiền.";

  if (order) {
    el.customerName.value = order.customerName || "";
    el.customerPhone.value = order.customerPhone || "";
    el.customerAddress.value = order.customerAddress || "";
    el.orderStatus.value = normalizeOrderStatus(order.status);
    el.orderTotalAmount.value = String(Number(order.totalAmount || 0));
  } else {
    resetOrderForm();
    el.orderStatus.value = ORDER_STATUS.PENDING;
  }

  renderSelectors();
  renderOrderHint();
  renderOrderDraft();

  if (preferredSku && state.products.some((product) => product.sku === preferredSku)) {
    selectOrderProduct(preferredSku);
    renderOrderHint();
  }

  requestAnimationFrame(() => {
    el.customerName.focus();
  });
}

function openOrderDetailModal(orderId) {
  const order = state.orders.find((entry) => entry.orderId === orderId);
  if (!order) {
    setStatus("Không tìm thấy đơn hàng.", "error");
    return;
  }

  if (state.activeModal && state.activeModal !== "order-detail") {
    closeActiveModal();
  }

  state.activeModal = "order-detail";
  el.orderDetailModal.hidden = false;
  el.orderDetailModal.setAttribute("aria-hidden", "false");
  document.body.classList.add("modal-open");
  renderOrderDetail(order);
}

function closeActiveModal() {
  if (state.activeModal === "product") {
    el.productModal.hidden = true;
    el.productModal.setAttribute("aria-hidden", "true");
  }

  if (state.activeModal === "product-detail") {
    el.productDetailModal.hidden = true;
    el.productDetailModal.setAttribute("aria-hidden", "true");
    el.productDetailBody.innerHTML = "";
  }

  if (state.activeModal === "stock") {
    el.stockModal.hidden = true;
    el.stockModal.setAttribute("aria-hidden", "true");
    state.stockTargetProductId = "";
    state.stockPreferredSku = "";
    state.stockVariants = [];
    renderStockModalProductCard();
    renderStockVariantPreview();
    renderStockVariantDraft();
    resetStockComposer();
  }

  if (state.activeModal === "order") {
    el.orderModal.hidden = true;
    el.orderModal.setAttribute("aria-hidden", "true");
    resetOrderForm();
    el.orderStatus.value = ORDER_STATUS.PENDING;
    state.orderModalMode = "create";
    state.editingOrderId = "";
    state.orderOriginalItems = [];
    renderOrderDraft();
    renderOrderHint();
  }

  if (state.activeModal === "order-detail") {
    el.orderDetailModal.hidden = true;
    el.orderDetailModal.setAttribute("aria-hidden", "true");
    el.orderDetailBody.innerHTML = "";
  }

  state.activeModal = "";
  document.body.classList.remove("modal-open");
}

async function handleImageChange(event) {
  const file = event.target.files[0];
  clearPendingImage();
  renderImagePreview();

  if (!file) {
    return;
  }

  setStatus("Đang nén ảnh ngay trên thiết bị trước khi tải lên...", "info");

  try {
    state.pendingImage = await compressImage(file);
    renderImagePreview();
    setStatus("Ảnh đã sẵn sàng. Việc thêm sản phẩm sẽ chỉ dùng một lần gọi API.", "success");
  } catch (error) {
    el.productImage.value = "";
    setStatus(error.message || "Không xử lý được ảnh đã chọn.", "error");
  }
}

function handleAddVariant() {
  const color = el.variantColor.value.trim();
  const size = el.variantSize.value.trim();
  const quantity = Number(el.variantQuantity.value);

  if (!color || !size) {
    setStatus("Vui lòng nhập màu và size cho biến thể.", "error");
    return;
  }

  if (!Number.isInteger(quantity) || quantity < 0) {
    setStatus("Số lượng nhập kho ban đầu phải từ 0 trở lên.", "error");
    return;
  }

  const duplicate = state.productVariants.find(
    (variant) => variant.color.toLowerCase() === color.toLowerCase() && variant.size.toLowerCase() === size.toLowerCase()
  );

  if (duplicate) {
    setStatus(`Biến thể ${color} · ${size} đã tồn tại trong danh sách.`, "error");
    return;
  }

  state.productVariants = state.productVariants
    .concat([{ color, size, quantity }])
    .sort((left, right) => left.color.localeCompare(right.color) || left.size.localeCompare(right.size));

  resetVariantComposer({ keepColor: true });
  renderVariantDraft();
  renderVariantPreview();
  el.variantSize.focus();
  setStatus(`Đã thêm biến thể ${color} · ${size}${quantity ? ` và nhập kho ${quantity}` : ""}.`, "success");
}

function handleVariantDraftActions(event) {
  const button = event.target.closest("[data-variant-action]");
  if (!button) {
    return;
  }

  const color = button.dataset.variantColor;
  const size = button.dataset.variantSize;

  state.productVariants = state.productVariants.filter(
    (variant) => !(variant.color === color && variant.size === size)
  );

  renderVariantDraft();
  renderVariantPreview();
  setStatus(`Đã xóa biến thể ${color} · ${size}.`, "info");
}

function syncStockVariantDefaults() {
  const existing = getExistingVariantByDraftInput();

  if (existing) {
    el.stockVariantHint.textContent = `${existing.sku} · Tồn hiện tại ${existing.currentStock}`;
    return;
  }
  el.stockVariantHint.textContent = "Chọn màu và size rồi thêm vào danh sách nhập kho.";
}

async function handleAddProduct(event) {
  event.preventDefault();

  if (state.isBusy) {
    return;
  }

  const payload = {
    name: el.productName.value.trim(),
    material: el.productMaterial.value.trim(),
    variants: state.productVariants.map((variant) => ({
      color: variant.color,
      size: variant.size,
      quantity: variant.quantity
    })),
    image: state.pendingImage
      ? {
          base64: state.pendingImage.base64,
          fileName: state.pendingImage.fileName,
          mimeType: state.pendingImage.mimeType
        }
      : null
  };

  if (!payload.name || !payload.material) {
    setStatus("Tên mẫu và chất liệu là bắt buộc.", "error");
    return;
  }

  if (!payload.variants.length) {
    setStatus("Hãy thêm ít nhất một biến thể màu và size trước khi lưu.", "error");
    return;
  }

  if (!payload.image) {
    setStatus("Vui lòng chọn ảnh sản phẩm trước khi lưu.", "error");
    return;
  }

  setBusy("product", true);

  try {
    const response = await postAction("addProductFull", payload);
    const variants = (response.variants || []).map(normalizeProduct);
    const createdSkus = new Set(variants.map((item) => item.sku));
    state.products = variants.concat(state.products.filter((item) => !createdSkus.has(item.sku)));
    state.summary = normalizeSummary(response.summary || buildSummaryFromState());
    state.lastSync = new Date().toISOString();
    resetProductForm();
    render();
    closeActiveModal();
    switchSection("products");
    setStatus(`Đã lưu mẫu ${response.product?.name || payload.name} với ${variants.length} biến thể.`, "success");
  } catch (error) {
    setStatus(error.message || "Không thể thêm sản phẩm.", "error");
  } finally {
    setBusy("product", false);
  }
}

async function handleStockIn(event) {
  event.preventDefault();

  if (state.isBusy) {
    return;
  }

  const productId = state.stockTargetProductId;
  const variants = state.stockVariants.map((variant) => ({
    color: variant.color,
    size: variant.size,
    quantity: variant.quantity
  }));

  if (!productId) {
    setStatus("Chưa chọn mẫu áo dài để nhập kho.", "error");
    return;
  }

  if (!variants.length) {
    setStatus("Hãy thêm ít nhất một dòng nhập kho trước khi lưu.", "error");
    return;
  }

  setBusy("stock", true);

  try {
    const response = await postAction("stockIn", {
      productId,
      variants
    });

    upsertProductsState((response.variants || []).map(normalizeProduct));
    state.summary = normalizeSummary(response.summary || buildSummaryFromState());
    state.lastSync = new Date().toISOString();
    render();
    closeActiveModal();
    switchSection("products");
    setStatus(`Đã cập nhật kho cho mẫu ${response.product?.name || productId}.`, "success");
  } catch (error) {
    setStatus(error.message || "Không thể cập nhật tồn kho.", "error");
  } finally {
    setBusy("stock", false);
  }
}

function handleAddStockVariant() {
  const color = el.stockColor.value.trim();
  const size = el.stockSize.value.trim();
  const quantity = Number(el.stockQuantity.value);

  if (!color || !size) {
    setStatus("Vui lòng nhập màu và size trong popup nhập kho.", "error");
    return;
  }

  if (!Number.isInteger(quantity) || quantity < 1) {
    setStatus("Số lượng nhập kho phải lớn hơn hoặc bằng 1.", "error");
    return;
  }

  const duplicate = state.stockVariants.find(
    (variant) => variant.color.toLowerCase() === color.toLowerCase() && variant.size.toLowerCase() === size.toLowerCase()
  );

  if (duplicate) {
    setStatus(`Dòng ${color} · ${size} đã có trong danh sách nhập kho.`, "error");
    return;
  }

  state.stockVariants = state.stockVariants
    .concat([{ color, size, quantity }])
    .sort((left, right) => left.color.localeCompare(right.color) || left.size.localeCompare(right.size));

  resetStockComposer({ keepColor: true });
  renderStockVariantPreview();
  renderStockVariantDraft();
  el.stockSize.focus();
  setStatus(`Đã thêm dòng nhập kho ${color} · ${size}.`, "success");
}

function handleStockVariantDraftActions(event) {
  const button = event.target.closest("[data-stock-variant-action]");
  if (!button) {
    return;
  }

  const color = button.dataset.stockVariantColor;
  const size = button.dataset.stockVariantSize;

  state.stockVariants = state.stockVariants.filter(
    (variant) => !(variant.color === color && variant.size === size)
  );

  renderStockVariantPreview();
  renderStockVariantDraft();
  setStatus(`Đã xóa dòng nhập kho ${color} · ${size}.`, "info");
}

function handleAddOrderItem(event) {
  event.preventDefault();

  const product = getSelectedProduct(state.selectedOrderSku);
  const quantity = Number(el.orderQuantity.value);

  if (!product) {
    setStatus("Hãy chọn biến thể áo trước khi thêm vào đơn hàng.", "error");
    return;
  }

  if (!Number.isInteger(quantity) || quantity < 1) {
    setStatus("Số lượng đặt hàng phải lớn hơn hoặc bằng 1.", "error");
    return;
  }

  const existing = state.orderDraft.find((item) => item.sku === product.sku);
  const nextQuantity = (existing ? existing.quantity : 0) + quantity;
  const available = getAvailableStockForOrderSku(product.sku);

  if (nextQuantity > available) {
    setStatus(`Chỉ còn ${available} áo khả dụng cho ${product.sku}.`, "error");
    return;
  }

  if (existing) {
    existing.quantity = nextQuantity;
  } else {
    state.orderDraft.push({
      sku: product.sku,
      name: product.name,
      color: product.color,
      size: product.size,
      quantity
    });
  }

  el.orderQuantity.value = "1";
  renderOrderDraft();
  setStatus(`Đã thêm ${product.sku} vào nháp đơn hàng.`, "success");
}

function handleDraftActions(event) {
  const button = event.target.closest("[data-action]");
  if (!button) {
    return;
  }

  const sku = button.dataset.sku;
  const action = button.dataset.action;
  const item = state.orderDraft.find((entry) => entry.sku === sku);

  if (!item) {
    return;
  }

  const product = state.products.find((entry) => entry.sku === sku);
  const available = product ? getAvailableStockForOrderSku(sku) : 0;

  if (action === "inc") {
    if (item.quantity >= available) {
      setStatus(`Chỉ còn ${available} áo khả dụng cho ${sku}.`, "error");
      return;
    }
    item.quantity += 1;
  }

  if (action === "dec") {
    item.quantity -= 1;
    if (item.quantity <= 0) {
      state.orderDraft = state.orderDraft.filter((entry) => entry.sku !== sku);
    }
  }

  if (action === "remove") {
    state.orderDraft = state.orderDraft.filter((entry) => entry.sku !== sku);
  }

  renderOrderDraft();
}

async function handleSubmitOrder() {
  if (state.isBusy) {
    return;
  }

  const customerName = el.customerName.value.trim();
  const customerPhone = normalizePhone(el.customerPhone.value);
  const customerAddress = el.customerAddress.value.trim();
  const totalAmount = Number(el.orderTotalAmount.value);

  if (!customerName) {
    setStatus("Vui lòng nhập tên khách hàng.", "error");
    el.customerName.focus();
    return;
  }

  if (customerPhone.length < 8) {
    setStatus("Số điện thoại khách hàng không hợp lệ.", "error");
    el.customerPhone.focus();
    return;
  }

  if (!customerAddress) {
    setStatus("Vui lòng nhập địa chỉ khách hàng.", "error");
    el.customerAddress.focus();
    return;
  }

  if (!state.orderDraft.length) {
    setStatus("Hãy thêm ít nhất một dòng vào nháp đơn hàng.", "error");
    return;
  }

  const invalidItem = state.orderDraft.find((item) => {
    const product = state.products.find((entry) => entry.sku === item.sku);
    return !product || item.quantity > getAvailableStockForOrderSku(item.sku);
  });

  if (invalidItem) {
    setStatus(`Tồn kho không đủ cho ${invalidItem.sku}.`, "error");
    return;
  }

  if (!Number.isFinite(totalAmount) || totalAmount < 0) {
    setStatus("Tổng tiền đơn hàng không hợp lệ.", "error");
    return;
  }

  setBusy("order", true);

  try {
    const mode = state.orderModalMode;
    const payload = {
      customerName,
      customerPhone,
      customerAddress,
      status: el.orderStatus.value,
      totalAmount,
      items: state.orderDraft.map((item) => ({
        sku: item.sku,
        size: item.size,
        quantity: item.quantity
      }))
    };

    const action = state.orderModalMode === "edit" ? "updateOrder" : "createOrder";
    const response = await postAction(action, state.orderModalMode === "edit"
      ? { orderId: state.editingOrderId, ...payload }
      : payload);

    (response.updatedStocks || []).forEach(applyUpdatedStock);
    const order = normalizeOrder(response.order || {});
    state.orders = [order].concat(state.orders.filter((entry) => entry.orderId !== order.orderId));
    state.ordersLoaded = true;
    resetVisibleOrdersCount();
    state.summary = normalizeSummary(response.summary || buildSummaryFromState());
    state.lastSync = new Date().toISOString();
    render();
    closeActiveModal();
    switchSection("orders");
    setStatus(
      mode === "edit"
        ? `Đã cập nhật đơn hàng ${response.order.orderId} thành công.`
        : `Đã tạo đơn hàng ${response.order.orderId} thành công.`,
      "success"
    );
  } catch (error) {
    setStatus(error.message || "Không thể lưu đơn hàng.", "error");
  } finally {
    setBusy("order", false);
  }
}

async function handleOrdersListClick(event) {
  const loadMoreButton = event.target.closest("[data-orders-load-more]");
  if (loadMoreButton) {
    handleOrdersLoadMore();
    return;
  }

  const button = event.target.closest("[data-order-action]");
  if (!button || state.isBusy) {
    return;
  }

  const orderId = button.dataset.orderId;
  const action = button.dataset.orderAction;
  const currentOrder = state.orders.find((entry) => entry.orderId === orderId);

  if (!currentOrder) {
    return;
  }

  if (action === "detail") {
    openOrderDetailModal(orderId);
    return;
  }

  if (action === "edit") {
    if (state.activeModal === "order-detail") {
      closeActiveModal();
    }
    openOrderModal({ mode: "edit", orderId });
    return;
  }

  if (action !== "delete") {
    return;
  }

  if (!window.confirm(`Xóa đơn hàng ${orderId}? Tồn kho sẽ được hoàn lại.`)) {
    return;
  }

  setBusy("order-delete", true);

  try {
    const response = await postAction("deleteOrder", {
      orderId
    });

    (response.updatedStocks || []).forEach(applyUpdatedStock);
    state.orders = state.orders.filter((entry) => entry.orderId !== orderId);
    state.ordersLoaded = true;
    resetVisibleOrdersCount();
    state.summary = normalizeSummary(response.summary || buildSummaryFromState());
    state.lastSync = new Date().toISOString();
    render();
    if (state.activeModal === "order-detail") {
      closeActiveModal();
    }
    setStatus(`Đã xóa đơn hàng ${orderId}.`, "success");
  } catch (error) {
    setStatus(error.message || "Không thể xóa đơn hàng.", "error");
  } finally {
    setBusy("order-delete", false);
  }
}

function handleProductCardClick(event) {
  const actionButton = event.target.closest("[data-product-action]");
  if (!actionButton) {
    return;
  }

  const productId = actionButton.dataset.productId;
  if (!productId) {
    return;
  }

  const action = actionButton.dataset.productAction;
  if (action === "detail") {
    openProductDetailModal(productId);
    return;
  }

  if (action === "stock") {
    openStockModal(productId);
    return;
  }

  focusProductGroupFlow(productId, "orders");
}

function handleProductDetailClick(event) {
  const productButton = event.target.closest("[data-product-action]");
  if (productButton) {
    const productId = productButton.dataset.productId;
    const action = productButton.dataset.productAction;

    if (!productId) {
      return;
    }

    if (action === "stock") {
      openStockModal(productId);
      return;
    }

    if (action === "order") {
      focusProductGroupFlow(productId, "orders");
    }

    return;
  }

  const variantButton = event.target.closest("[data-variant-action]");
  if (!variantButton) {
    return;
  }

  const sku = variantButton.dataset.sku;
  if (!sku) {
    return;
  }

  focusProductFlow(sku, variantButton.dataset.variantAction === "stock" ? "inventory" : "orders");
}

function handleInventoryListClick(event) {
  const button = event.target.closest("[data-inventory-action]");
  if (!button) {
    return;
  }

  focusProductFlow(button.dataset.sku, button.dataset.inventoryAction === "stock" ? "inventory" : "orders");
}

function handleLowStockListClick(event) {
  const button = event.target.closest("[data-low-stock-sku]");
  if (!button) {
    return;
  }

  focusProductFlow(button.dataset.lowStockSku, "inventory");
}

function focusProductFlow(sku, sectionKey) {
  const product = getSelectedProduct(sku);
  if (!product) {
    return;
  }

  if (sectionKey === "inventory") {
    openStockModal(product.productId, sku);
    return;
  }

  openOrderModal({ mode: "create", preferredSku: sku });
  switchSection(sectionKey);
  setStatus(`Đã chọn ${sku} trong mục ${SECTION_META[sectionKey].title}.`, "info");
}

function focusProductGroupFlow(productId, sectionKey) {
  const variants = getVariantsByProductId(productId);
  if (!variants.length) {
    return;
  }

  const preferredVariant = sectionKey === "orders"
    ? variants.find((item) => item.currentStock > 0) || variants[0]
    : variants[0];

  focusProductFlow(preferredVariant.sku, sectionKey);

  if (sectionKey === "orders" && variants.length > 1) {
    setStatus(
      `Đã mở ${SECTION_META[sectionKey].title} cho mẫu ${preferredVariant.name}. Hãy chọn đúng màu và size trong danh sách biến thể.`,
      "info"
    );
  }
}

function render() {
  renderSectionState();
  renderSummary();
  renderImagePreview();
  renderVariantPreview();
  renderVariantDraft();
  renderSelectors();
  renderStockBadge();
  renderStockModalProductCard(state.stockPreferredSku);
  renderStockVariantPreview();
  renderStockVariantDraft();
  renderOrderHint();
  renderOrderDraft();
  renderProducts();
  renderInventoryList();
  renderDashboardLists();
  renderOrdersTabs();
  renderOrdersList();
  renderSyncNote();
}

function renderSectionState() {
  const meta = SECTION_META[state.activeSection];
  el.sectionTitle.textContent = meta.title;
  el.sectionDescription.textContent = meta.description;

  el.navButtons.forEach((button) => {
    button.classList.toggle("is-active", button.dataset.sectionTarget === state.activeSection);
  });

  el.sectionPanels.forEach((panel) => {
    panel.classList.toggle("is-active", panel.dataset.sectionPanel === state.activeSection);
  });
}

function renderSummary() {
  const summary = state.summary.productCount ? state.summary : buildSummaryFromState();
  el.totalProducts.textContent = String(summary.productCount || 0);
  el.totalStock.textContent = String(summary.totalStock || 0);
  el.totalOrders.textContent = String(summary.orderCount || 0);
}

function renderDashboardLists() {
  renderLowStockList();
  renderRecentList();
}

function renderLowStockList() {
  const lowStock = state.products
    .filter((product) => product.currentStock <= 3)
    .sort((left, right) => left.currentStock - right.currentStock || left.name.localeCompare(right.name));

  if (!lowStock.length) {
    el.lowStockList.innerHTML = renderEmptyState(
      "Hiện chưa có mẫu nào cần chú ý.",
      "Các biến thể sắp hết hàng sẽ tự động xuất hiện tại đây."
    );
    return;
  }

  el.lowStockList.innerHTML = lowStock
    .map((product) => {
      const stockClass = getStockClass(product.currentStock);
      return `
        <article class="entity-row">
          <div class="entity-mainline">
            <div class="entity-text">
              <strong>${escapeHtml(product.name)}</strong>
              <p>${escapeHtml(product.sku)} · ${escapeHtml(product.color)} · ${escapeHtml(product.size)}</p>
            </div>
            <span class="${stockClass}">Còn ${product.currentStock}</span>
          </div>
          <div class="entity-actions">
            <button type="button" class="small-button accent" data-low-stock-sku="${escapeHtml(product.sku)}">Nhập thêm</button>
          </div>
        </article>
      `;
    })
    .join("");
}

function renderRecentList() {
  const topSelling = getTopSellingProducts().slice(0, 5);

  if (!topSelling.length) {
    el.dashboardRecentList.innerHTML = renderEmptyState(
      "Chưa có dữ liệu bán chạy.",
      "Khi có đơn hàng, mẫu áo bán tốt sẽ xuất hiện tại đây."
    );
    return;
  }

  el.dashboardRecentList.innerHTML = topSelling
    .map(
      ({ product, soldUnits }) => `
        <article class="entity-row">
          <div class="entity-media">
            <img src="${escapeHtml(product.image)}" alt="${escapeHtml(product.name)}" class="entity-thumb">
            <div class="entity-text">
              <strong>${escapeHtml(product.name)}</strong>
              <p>${escapeHtml(product.material || "Chưa có chất liệu")} · ${product.colors.length} màu · ${product.sizes.length} size</p>
              <p>Đã bán ${soldUnits} · Tổng tồn ${product.totalStock}</p>
            </div>
          </div>
        </article>
      `
    )
    .join("");
}

function getTopSellingProducts() {
  if (!state.orders.length) {
    return [];
  }

  const catalogProducts = getCatalogProducts();
  const catalogById = new Map(catalogProducts.map((product) => [product.productId, product]));
  const variantToGroupId = new Map(state.products.map((product) => [product.sku, getProductGroupKey(product)]));
  const soldUnitsByProductId = new Map();

  state.orders.forEach((order) => {
    (order.items || []).forEach((item) => {
      const productId = variantToGroupId.get(item.sku);
      if (!productId) {
        return;
      }

      soldUnitsByProductId.set(
        productId,
        (soldUnitsByProductId.get(productId) || 0) + Number(item.quantity || 0)
      );
    });
  });

  return Array.from(soldUnitsByProductId.entries())
    .map(([productId, soldUnits]) => ({
      product: catalogById.get(productId),
      soldUnits
    }))
    .filter((entry) => entry.product && entry.soldUnits > 0)
    .sort((a, b) => {
      if (b.soldUnits !== a.soldUnits) {
        return b.soldUnits - a.soldUnits;
      }

      return a.product.name.localeCompare(b.product.name, "vi");
    });
}

function renderImagePreview() {
  if (!state.pendingImage) {
    el.imagePreview.className = "preview-card empty";
    el.imagePreview.innerHTML = `
      <div class="preview-placeholder">
        <strong>Ảnh sẽ được nén trước khi tải lên.</strong>
        <span>Chiều ngang tối đa 500px, chất lượng 0.7, mục tiêu dưới 300KB</span>
      </div>
    `;
    return;
  }

  el.imagePreview.className = "preview-card";
  el.imagePreview.innerHTML = `
    <div class="preview-media">
      <img src="${escapeHtml(state.pendingImage.previewUrl)}" alt="Xem trước">
      <div class="preview-meta">
        <strong>${escapeHtml(state.pendingImage.fileName)}</strong>
        <span>${state.pendingImage.width} × ${state.pendingImage.height}px</span>
        <span>${humanFileSize(state.pendingImage.bytes)}</span>
        <span>${state.pendingImage.bytes <= CONFIG.TARGET_IMAGE_BYTES ? "Đã đạt kích thước mục tiêu." : "Đã nén tối đa có thể."}</span>
      </div>
    </div>
  `;
}

function renderVariantPreview() {
  if (!state.productVariants.length) {
    el.variantPreview.className = "helper-text variant-preview";
    el.variantPreview.innerHTML = `
      <strong>Chưa có biến thể nào.</strong>
      <p>Nhập màu, size và có thể nhập tồn kho ban đầu ngay khi thêm biến thể.</p>
    `;
    renderProductDraftMeta();
    return;
  }

  const colors = Array.from(new Set(state.productVariants.map((variant) => variant.color)));
  const sizes = Array.from(new Set(state.productVariants.map((variant) => variant.size)));
  const totalInitialStock = state.productVariants.reduce((sum, variant) => sum + Number(variant.quantity || 0), 0);
  el.variantPreview.className = "helper-text variant-preview";
  el.variantPreview.innerHTML = `
    <strong>${state.productVariants.length} biến thể đã sẵn sàng để lưu</strong>
    <p>Màu: ${escapeHtml(colors.join(", "))}</p>
    <p>Size: ${escapeHtml(sizes.join(", "))}</p>
    <p>Tồn kho nhập ngay: ${totalInitialStock}</p>
  `;
  renderProductDraftMeta();
}

function renderVariantDraft() {
  if (!state.productVariants.length) {
    el.variantDraftList.innerHTML = renderEmptyState(
      "Danh sách biến thể đang trống.",
      "Thêm từng dòng màu và size để chuẩn bị lưu mẫu áo dài."
    );
    return;
  }

  el.variantDraftList.innerHTML = state.productVariants
    .map((variant) => `
      <article class="variant-row">
        <div class="variant-row-top">
            <div class="entity-text">
              <strong>${escapeHtml(variant.color)} · ${escapeHtml(variant.size)}</strong>
              <p>Nhập kho ngay: ${variant.quantity}</p>
            </div>
          <button
            type="button"
            class="small-button danger"
            data-variant-action="remove"
            data-variant-color="${escapeHtml(variant.color)}"
            data-variant-size="${escapeHtml(variant.size)}"
          >
            Xóa
          </button>
        </div>
      </article>
    `)
    .join("");
}

function renderSelectors() {
  if (!state.products.length) {
    el.orderProductResults.innerHTML = renderCompactOrderSearchEmptyState(
      "Chưa có biến thể áo.",
      "Thêm mẫu và nhập kho trước khi tạo đơn."
    );
    return;
  }

  renderOrderProductResults();
}

function renderStockBadge() {
  const catalogProduct = getCatalogProducts().find((product) => product.productId === state.stockTargetProductId);
  const totalStock = catalogProduct ? catalogProduct.totalStock : 0;
  el.stockCurrentBadge.textContent = `Tổng tồn hiện tại: ${totalStock}`;
}

function renderOrderHint() {
  const product = getSelectedProduct(state.selectedOrderSku);
  if (!product) {
    el.orderHint.textContent = "Tìm biến thể áo rồi chạm để chọn trước khi thêm vào nháp đơn hàng.";
    return;
  }

  el.orderHint.textContent = `${product.sku} · ${product.color} · ${product.size} · Khả dụng ${getAvailableStockForOrderSku(product.sku)}`;
}

function handleOrderProductSearch() {
  if (!state.selectedOrderSku) {
    renderOrderProductResults();
    renderOrderHint();
    return;
  }

  const selectedProduct = getSelectedProduct(state.selectedOrderSku);
  const normalizedSearch = normalizeText(el.orderProductSearch.value);
  const normalizedSelected = selectedProduct ? normalizeText(buildOrderProductSearchLabel(selectedProduct)) : "";

  if (!selectedProduct || (normalizedSearch && normalizedSearch !== normalizedSelected)) {
    state.selectedOrderSku = "";
  }

  renderOrderProductResults();
  renderOrderHint();
}

function handleOrderProductResultsClick(event) {
  const button = event.target.closest("[data-order-product-sku]");
  if (!button) {
    return;
  }

  selectOrderProduct(button.dataset.orderProductSku);
  renderOrderHint();
}

function selectOrderProduct(sku) {
  const product = getSelectedProduct(sku);
  if (!product) {
    return;
  }

  state.selectedOrderSku = product.sku;
  el.orderProductSearch.value = buildOrderProductSearchLabel(product);
  renderOrderProductResults();
}

function renderOrderProductResults() {
  const matches = getFilteredOrderProducts();

  if (!matches.length) {
    el.orderProductResults.innerHTML = renderCompactOrderSearchEmptyState(
      "Không tìm thấy biến thể phù hợp.",
      "Thử tên áo, màu, size hoặc mã mẫu khác."
    );
    return;
  }

  el.orderProductResults.innerHTML = matches
    .slice(0, 12)
    .map((product) => {
      const isSelected = product.sku === state.selectedOrderSku;
      return `
        <button
          type="button"
          class="order-product-option${isSelected ? " is-selected" : ""}"
          data-order-product-sku="${escapeHtml(product.sku)}"
        >
          <strong>${escapeHtml(product.name)} · ${escapeHtml(product.color)} · ${escapeHtml(product.size)}</strong>
          <span>${escapeHtml(product.productId)} · Tồn ${product.currentStock}</span>
        </button>
      `;
    })
    .join("");
}

function getFilteredOrderProducts() {
  const query = normalizeText(el.orderProductSearch.value);

  return state.products
    .slice()
    .filter((product) => {
      if (!query) {
        return product.currentStock > 0 || product.sku === state.selectedOrderSku;
      }

      const haystack = normalizeText([
        product.name,
        product.productId,
        product.sku,
        product.material,
        product.color,
        product.size
      ].join(" "));

      return haystack.includes(query);
    })
    .sort((left, right) => {
      const availabilityDelta = Number(right.currentStock > 0) - Number(left.currentStock > 0);
      if (availabilityDelta !== 0) {
        return availabilityDelta;
      }

      return left.name.localeCompare(right.name, "vi")
        || left.color.localeCompare(right.color, "vi")
        || left.size.localeCompare(right.size, "vi");
    });
}

function buildOrderProductSearchLabel(product) {
  return `${product.name} · ${product.color} · ${product.size}`;
}

function renderCompactOrderSearchEmptyState(title, message) {
  return `
    <div class="order-product-empty">
      <strong>${escapeHtml(title)}</strong>
      <span>${escapeHtml(message)}</span>
    </div>
  `;
}

function renderOrderDraft() {
  if (!state.orderDraft.length) {
    el.orderDraft.innerHTML = renderEmptyState(
      "Nháp đơn hàng đang trống.",
      "Chọn một biến thể rồi thêm nó vào nháp đơn hàng."
    );
    el.orderDraftTotal.textContent = "0 áo";
    el.orderDraftSummary.textContent = "Chưa có dòng nào trong nháp.";
    return;
  }

  el.orderDraft.innerHTML = state.orderDraft
    .map((item) => {
      return `
        <article class="draft-item">
          <div class="draft-top">
            <div class="entity-text">
              <strong>${escapeHtml(item.name)}</strong>
              <p>${escapeHtml(item.sku)} · ${escapeHtml(item.color)} · ${escapeHtml(item.size)}</p>
            </div>
            <strong>${item.quantity} áo</strong>
          </div>
          <div class="draft-top">
            <p>Màu ${escapeHtml(item.color)} · Size ${escapeHtml(item.size)}</p>
            <div class="entity-actions">
              <button type="button" class="small-button" data-action="dec" data-sku="${escapeHtml(item.sku)}" aria-label="Giảm số lượng">-</button>
              <button type="button" class="small-button" data-action="inc" data-sku="${escapeHtml(item.sku)}" aria-label="Tăng số lượng">+</button>
              <button type="button" class="small-button danger" data-action="remove" data-sku="${escapeHtml(item.sku)}" aria-label="Xóa dòng">Xóa</button>
            </div>
          </div>
        </article>
      `;
    })
    .join("");

  const totalUnits = state.orderDraft.reduce((sum, item) => sum + item.quantity, 0);
  el.orderDraftTotal.textContent = `${totalUnits} áo`;
  el.orderDraftSummary.textContent = `Tổng tiền đơn sẽ nhập thủ công bên trên.`;
}

function renderProducts() {
  const products = getVisibleProducts();

  if (!products.length) {
    el.productList.innerHTML = renderEmptyState(
      state.products.length ? "Không có sản phẩm phù hợp với từ khóa." : "Chưa có sản phẩm nào.",
      state.products.length ? "Thử từ khóa khác hoặc xóa nội dung tìm kiếm." : "Thêm sản phẩm đầu tiên để bắt đầu xây dựng danh mục."
    );
    return;
  }

  el.productList.innerHTML = products
    .map((product) => {
      const stockClass = getStockClass(product.totalStock);
      return `
        <article class="product-card" data-product-id="${escapeHtml(product.productId)}">
          <img src="${escapeHtml(product.image)}" alt="${escapeHtml(product.name)}" loading="lazy" decoding="async">
          <div class="product-title">
            <h4>${escapeHtml(product.name)}</h4>
            <p>${escapeHtml(product.material || "Chưa có chất liệu")}</p>
          </div>
          <span class="${stockClass}">Tổng tồn ${product.totalStock}</span>
          <div class="product-actions">
            <button type="button" class="small-button" data-product-action="detail" data-product-id="${escapeHtml(product.productId)}">Chi tiết</button>
            <button type="button" class="small-button accent" data-product-action="stock" data-product-id="${escapeHtml(product.productId)}">Nhập kho</button>
            <button type="button" class="small-button success" data-product-action="order" data-product-id="${escapeHtml(product.productId)}">Bán</button>
          </div>
        </article>
      `;
    })
    .join("");

}

function renderProductDetail(product) {
  const stockClass = getStockClass(product.totalStock);
  const variantGroups = getProductVariantGroups(product);
  const variantMarkup = variantGroups.map((group) => {
    const groupStock = group.variants.reduce((sum, variant) => sum + Number(variant.currentStock || 0), 0);
    return `
      <section class="variant-group">
        <div class="variant-group-head">
          <div class="entity-text">
            <strong>Màu ${escapeHtml(group.color)}</strong>
            <p>${group.variants.length} size · Tổng tồn ${groupStock}</p>
          </div>
          <span class="${getStockClass(groupStock)}">Tồn ${groupStock}</span>
        </div>

        <div class="variant-group-list">
          ${group.variants.map((variant) => `
            <article class="variant-detail-row">
              <div class="variant-detail-top">
                <div class="entity-text">
                  <strong>Size ${escapeHtml(variant.size)}</strong>
                  <p>${escapeHtml(variant.sku)}</p>
                </div>
                <div class="entity-text product-detail-price">
                  <strong>Tồn ${variant.currentStock}</strong>
                  <p>Size ${escapeHtml(variant.size)}</p>
                </div>
              </div>
              <div class="variant-detail-actions">
                <button
                  type="button"
                  class="small-button accent"
                  data-variant-action="stock"
                  data-sku="${escapeHtml(variant.sku)}"
                >
                  Nhập kho
                </button>
                <button
                  type="button"
                  class="small-button success"
                  data-variant-action="order"
                  data-sku="${escapeHtml(variant.sku)}"
                >
                  Bán
                </button>
              </div>
            </article>
          `).join("")}
        </div>
      </section>
    `;
  }).join("");

  el.productDetailTitle.textContent = product.name || "Chi tiết sản phẩm";
  el.productDetailBody.innerHTML = `
    <section class="product-detail-summary">
      <img class="product-detail-image" src="${escapeHtml(product.image)}" alt="${escapeHtml(product.name)}" loading="lazy" decoding="async">
      <div class="product-detail-copy">
        <strong>${escapeHtml(product.name)}</strong>
        <p>${escapeHtml(product.material || "Chưa có chất liệu")}</p>
        <p>Mã mẫu: ${escapeHtml(product.productId)}</p>
        <span class="${stockClass}">Tổng tồn ${product.totalStock}</span>
      </div>
    </section>

    <section class="order-detail-total product-detail-total">
      <span>Biến thể</span>
      <strong>${product.variantCount} biến thể · ${product.colors.length} màu · ${product.sizes.length} size</strong>
    </section>

    <div class="product-detail-actions">
      <button type="button" class="small-button accent" data-product-action="stock" data-product-id="${escapeHtml(product.productId)}">Nhập kho</button>
      <button type="button" class="small-button success" data-product-action="order" data-product-id="${escapeHtml(product.productId)}">Bán</button>
    </div>

    <section class="product-detail-meta">
      ${variantMarkup || renderEmptyState("Chưa có biến thể.", "Bấm Nhập kho để thêm màu, size và số lượng.")}
    </section>
  `;
}

function renderStockModalProductCard(preferredSku = "") {
  if (!el.stockModalProductCard) {
    return;
  }

  const catalogProduct = getCatalogProducts().find((product) => product.productId === state.stockTargetProductId);
  if (!catalogProduct) {
    el.stockModalProductCard.innerHTML = `
      <div class="empty-state compact-empty-state">
        <div>
          <strong>Chưa chọn mẫu áo dài.</strong>
          <p>Hãy bấm nút Nhập kho trên một card sản phẩm để mở popup này.</p>
        </div>
      </div>
    `;
    return;
  }

  const preferredVariant = preferredSku ? getSelectedProduct(preferredSku) : null;
  const colorsLine = catalogProduct.colors.length ? catalogProduct.colors.join(", ") : "Chưa có màu";
  const sizesLine = catalogProduct.sizes.length ? catalogProduct.sizes.join(", ") : "Chưa có size";
  const preferredCopy = preferredVariant
    ? `<p>Gợi ý theo biến thể đang chọn: ${escapeHtml(preferredVariant.color)} · ${escapeHtml(preferredVariant.size)} · Tồn ${preferredVariant.currentStock}</p>`
    : "";

  el.stockModalProductCard.innerHTML = `
    <div class="product-summary-media">
      <img src="${escapeHtml(catalogProduct.image)}" alt="${escapeHtml(catalogProduct.name)}" loading="lazy" decoding="async">
      <div class="product-summary-copy">
        <strong>${escapeHtml(catalogProduct.name)}</strong>
        <p>${escapeHtml(catalogProduct.material || "Chưa có chất liệu")}</p>
        <p>Tổng tồn ${catalogProduct.totalStock}</p>
        <p>Màu: ${escapeHtml(colorsLine)}</p>
        <p>Size: ${escapeHtml(sizesLine)}</p>
        ${preferredCopy}
      </div>
    </div>
  `;
}

function renderStockVariantPreview() {
  if (!state.stockVariants.length) {
    el.stockVariantPreview.className = "helper-text variant-preview";
    el.stockVariantPreview.innerHTML = `
      <strong>Chưa có dòng nhập kho nào.</strong>
      <p>Thêm màu, size và số lượng để chuẩn bị lưu vào kho.</p>
    `;
    renderStockDraftMeta();
    return;
  }

  const colors = Array.from(new Set(state.stockVariants.map((variant) => variant.color)));
  const sizes = Array.from(new Set(state.stockVariants.map((variant) => variant.size)));
  const totalQuantity = state.stockVariants.reduce((sum, variant) => sum + Number(variant.quantity || 0), 0);
  el.stockVariantPreview.className = "helper-text variant-preview";
  el.stockVariantPreview.innerHTML = `
    <strong>${state.stockVariants.length} dòng nhập kho đã sẵn sàng</strong>
    <p>Màu: ${escapeHtml(colors.join(", "))}</p>
    <p>Size: ${escapeHtml(sizes.join(", "))}</p>
    <p>Tổng số lượng nhập: ${totalQuantity}</p>
  `;
  renderStockDraftMeta();
}

function renderStockVariantDraft() {
  if (!state.stockVariants.length) {
    el.stockVariantDraftList.innerHTML = renderEmptyState(
      "Danh sách nhập kho đang trống.",
      "Thêm từng dòng màu, size và số lượng để lưu kho."
    );
    return;
  }

  el.stockVariantDraftList.innerHTML = state.stockVariants
    .map((variant) => `
      <article class="variant-row">
        <div class="variant-row-top">
          <div class="entity-text">
            <strong>${escapeHtml(variant.color)} · ${escapeHtml(variant.size)}</strong>
            <p>Số lượng nhập: ${variant.quantity}</p>
          </div>
          <button
            type="button"
            class="small-button danger"
            data-stock-variant-action="remove"
            data-stock-variant-color="${escapeHtml(variant.color)}"
            data-stock-variant-size="${escapeHtml(variant.size)}"
          >
            Xóa
          </button>
        </div>
      </article>
    `)
    .join("");
}

function renderInventoryList() {
  if (!state.products.length) {
    el.inventoryList.innerHTML = renderEmptyState(
      "Chưa có dữ liệu tồn kho.",
      "Sản phẩm và số lượng tồn sẽ xuất hiện ở đây sau lần lưu đầu tiên."
    );
    return;
  }

  const inventory = state.products
    .slice()
    .sort((left, right) => left.currentStock - right.currentStock || left.name.localeCompare(right.name));

  el.inventoryList.innerHTML = inventory
    .map((product) => {
      const stockClass = getStockClass(product.currentStock);
      return `
        <article class="entity-row">
          <div class="entity-mainline">
            <div class="entity-text">
              <strong>${escapeHtml(product.name)}</strong>
              <p>${escapeHtml(product.sku)} · ${escapeHtml(product.color)} · ${escapeHtml(product.size)}</p>
            </div>
            <span class="${stockClass}">Còn ${product.currentStock}</span>
          </div>
          <div class="entity-actions">
            <button type="button" class="small-button accent" data-inventory-action="stock" data-sku="${escapeHtml(product.sku)}">Nhập thêm</button>
            <button type="button" class="small-button success" data-inventory-action="order" data-sku="${escapeHtml(product.sku)}">Bán</button>
          </div>
        </article>
      `;
    })
    .join("");
}

function renderOrdersList() {
  const filteredOrders = state.orders.filter((order) =>
    state.orderFilter === "sent"
      ? order.status === ORDER_STATUS.SENT
      : order.status !== ORDER_STATUS.SENT
  );
  const visibleOrders = filteredOrders.slice(0, state.visibleOrdersCount);

  if (state.isLoadingOrders && !state.orders.length) {
    renderOrdersLoading();
    return;
  }

  if (!state.ordersLoaded && !state.orders.length) {
    el.ordersList.innerHTML = renderEmptyState(
      "Chưa tải danh sách đơn.",
      "Danh sách đơn sẽ được tải khi bạn mở mục Đơn."
    );
    return;
  }

  if (!filteredOrders.length) {
    el.ordersList.innerHTML = renderEmptyState(
      state.orderFilter === "sent" ? "Chưa có đơn đã giao." : "Chưa có đơn chờ giao.",
      state.orderFilter === "sent"
        ? "Những đơn đã giao cho đơn vị vận chuyển sẽ hiện ở đây."
        : "Những đơn chưa giao cho đơn vị vận chuyển sẽ hiện ở đây."
    );
    return;
  }

  const ordersMarkup = visibleOrders
    .map((order) => {
      const statusMeta = getOrderStatusMeta(order.status);
      const totalUnits = getOrderTotalUnits(order);
      const totalValue = getOrderTotalValue(order);
      const customerName = order.customerName || "Khách lẻ";

      return `
        <article class="entity-row order-card">
          <div class="order-card-header">
            <div class="order-card-meta">
              <strong>${escapeHtml(customerName)}</strong>
              <p>${escapeHtml(formatDateTime(order.date))}</p>
            </div>
            <span class="status-chip compact ${statusMeta.className}" aria-label="${escapeHtml(statusMeta.label)}" title="${escapeHtml(statusMeta.label)}">
              <span class="status-chip-icon" aria-hidden="true">${statusMeta.icon}</span>
            </span>
          </div>

          <div class="order-summary-line compact">
            <div class="order-summary-cell">
              <span>SĐT</span>
              <strong>${escapeHtml(order.customerPhone || "Chưa có")}</strong>
            </div>
            <div class="order-summary-cell">
              <span>Số lượng</span>
              <strong>${totalUnits} áo</strong>
            </div>
            <div class="order-summary-cell order-summary-total">
              <span>Tổng tiền</span>
              <strong>${formatCurrency(totalValue)}</strong>
            </div>
          </div>

          <div class="order-card-actions">
            <button
              type="button"
              class="small-button accent"
              data-order-action="detail"
              data-order-id="${escapeHtml(order.orderId)}"
            >
              Chi tiết
            </button>
            <button
              type="button"
              class="small-button"
              data-order-action="edit"
              data-order-id="${escapeHtml(order.orderId)}"
            >
              Sửa
            </button>
            <button
              type="button"
              class="small-button danger"
              data-order-action="delete"
              data-order-id="${escapeHtml(order.orderId)}"
            >
              Xóa
            </button>
          </div>
        </article>
      `;
    })
    .join("");

  const hasMore = filteredOrders.length > visibleOrders.length;
  const loadMoreMarkup = hasMore
    ? `
      <button type="button" class="secondary-button orders-load-more" data-orders-load-more>
        Tải thêm ${Math.min(CONFIG.ORDERS_PAGE_SIZE, filteredOrders.length - visibleOrders.length)} đơn
      </button>
    `
    : "";

  el.ordersList.innerHTML = ordersMarkup + loadMoreMarkup;
}

function renderOrdersTabs() {
  if (!el.ordersStatusTabs) {
    return;
  }

  const buttons = Array.from(el.ordersStatusTabs.querySelectorAll("[data-order-filter]"));
  buttons.forEach((button) => {
    const isActive = button.dataset.orderFilter === state.orderFilter;
    button.classList.toggle("is-active", isActive);
    button.setAttribute("aria-selected", String(isActive));
  });
}

function handleOrdersLoadMore() {
  state.visibleOrdersCount += CONFIG.ORDERS_PAGE_SIZE;
  renderOrdersList();
}

function renderOrdersLoading() {
  el.ordersList.innerHTML = `
    <div class="skeleton-card"></div>
    <div class="skeleton-card"></div>
  `;
}

function renderOrderDetail(order) {
  const statusMeta = getOrderStatusMeta(order.status);
  const totalUnits = getOrderTotalUnits(order);
  const totalValue = getOrderTotalValue(order);
  const itemsMarkup = order.items.length
    ? order.items.map((item) => {
      return `
        <div class="order-item-line">
          <div class="entity-text">
            <strong>${escapeHtml(item.name || item.sku)}</strong>
            <p>${escapeHtml(item.sku)} · ${escapeHtml(item.color || "")} · ${escapeHtml(item.size)}</p>
          </div>
          <div class="entity-text order-summary-total">
            <strong>${item.quantity} áo</strong>
            <p>${escapeHtml(item.color || "Chưa có màu")} · ${escapeHtml(item.size || "Chưa có size")}</p>
          </div>
        </div>
      `;
    }).join("")
    : renderEmptyState("Đơn chưa có áo.", "Hãy bấm sửa để thêm các áo vào đơn này.");

  el.orderDetailTitle.textContent = order.orderId || "Chi tiết đơn hàng";
  el.orderDetailBody.innerHTML = `
    <section class="order-detail-hero">
      <div class="order-card-header order-detail-hero-top">
        <div class="order-card-meta order-detail-meta">
          <span class="order-detail-eyebrow">Mã đơn</span>
          <strong>${escapeHtml(order.orderId)}</strong>
          <p>${escapeHtml(formatDateTime(order.date))}</p>
        </div>
        <span class="status-chip ${statusMeta.className}">${escapeHtml(statusMeta.label)}</span>
      </div>

      <div class="order-detail-stats">
        <div class="order-detail-stat">
          <span>Tổng mẫu</span>
          <strong>${order.items.length}</strong>
        </div>
        <div class="order-detail-stat">
          <span>Số áo</span>
          <strong>${totalUnits} áo</strong>
        </div>
        <div class="order-detail-stat">
          <span>Tổng tiền đơn</span>
          <strong>${formatCurrency(totalValue)}</strong>
        </div>
      </div>
    </section>

    <section class="order-detail-section">
      <p class="panel-kicker">Khách hàng</p>
      <div class="order-customer order-detail-info-card">
        <div class="order-detail-info-row">
          <span>Người nhận</span>
          <strong>${escapeHtml(order.customerName || "Khách lẻ")}</strong>
        </div>
        <div class="order-detail-info-row">
          <span>Điện thoại</span>
          <strong>${escapeHtml(order.customerPhone || "Chưa có số điện thoại")}</strong>
        </div>
        <div class="order-detail-info-row">
          <span>Địa chỉ</span>
          <strong>${escapeHtml(order.customerAddress || "Chưa có địa chỉ")}</strong>
        </div>
      </div>
    </section>

    <section class="order-detail-section">
      <p class="panel-kicker">Sản phẩm</p>
      <div class="order-items-stack">${itemsMarkup}</div>
    </section>

    <section class="order-detail-total">
      <span>Tổng đơn</span>
      <strong>${totalUnits} áo · ${formatCurrency(totalValue)}</strong>
    </section>

    <div class="order-card-actions">
      <button
        type="button"
        class="small-button"
        data-order-action="edit"
        data-order-id="${escapeHtml(order.orderId)}"
      >
        Sửa đơn
      </button>
      <button
        type="button"
        class="small-button danger"
        data-order-action="delete"
        data-order-id="${escapeHtml(order.orderId)}"
      >
        Xóa đơn
      </button>
    </div>
  `;
}

function renderSkeletons() {
  el.productList.innerHTML = `
    <div class="skeleton-card"></div>
    <div class="skeleton-card"></div>
    <div class="skeleton-card"></div>
  `;
}

function renderSyncNote() {
  el.syncNote.textContent = state.lastSync ? `Cập nhật lúc ${formatTime(state.lastSync)}` : "Chưa đồng bộ";
}

async function postAction(action, payload = {}) {
  if (!CONFIG.API_URL || CONFIG.API_URL.includes("YOUR_DEPLOYMENT_ID")) {
    throw new Error("Hãy cập nhật CONFIG.API_URL trong app.js bằng URL Apps Script đã triển khai.");
  }

  const apiUrl = new URL(CONFIG.API_URL);
  apiUrl.searchParams.set("_ts", String(Date.now()));

  const response = await fetch(apiUrl.toString(), {
    method: "POST",
    cache: "no-store",
    headers: {
      "Content-Type": "text/plain;charset=utf-8"
    },
    body: JSON.stringify({
      action,
      ...payload
    })
  });

  const text = await response.text();
  let data;

  try {
    data = JSON.parse(text);
  } catch (error) {
    throw new Error("API không trả về JSON hợp lệ.");
  }

  if (!response.ok || data.error) {
    throw new Error(data.error || "Gọi API thất bại.");
  }

  return data;
}

async function compressImage(file) {
  if (!file.type.startsWith("image/")) {
    throw new Error("Tệp được chọn không phải là ảnh.");
  }

  const image = await loadImage(file);
  const sourceWidth = image.naturalWidth || image.width;
  const sourceHeight = image.naturalHeight || image.height;
  let targetWidth = Math.min(CONFIG.MAX_IMAGE_WIDTH, sourceWidth);
  let targetHeight = Math.round(sourceHeight * (targetWidth / sourceWidth));
  let blob = await drawCompressedBlob(image, targetWidth, targetHeight, CONFIG.IMAGE_QUALITY);

  while (blob.size > CONFIG.TARGET_IMAGE_BYTES && targetWidth > 220) {
    targetWidth = Math.round(targetWidth * 0.9);
    targetHeight = Math.round(sourceHeight * (targetWidth / sourceWidth));
    blob = await drawCompressedBlob(image, targetWidth, targetHeight, CONFIG.IMAGE_QUALITY);
  }

  return {
    fileName: buildUploadFileName(file.name),
    mimeType: "image/jpeg",
    width: targetWidth,
    height: targetHeight,
    bytes: blob.size,
    previewUrl: URL.createObjectURL(blob),
    base64: await blobToBase64(blob)
  };
}

function loadImage(file) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    const objectUrl = URL.createObjectURL(file);

    image.onload = () => {
      URL.revokeObjectURL(objectUrl);
      resolve(image);
    };

    image.onerror = () => {
      URL.revokeObjectURL(objectUrl);
      reject(new Error("Không xử lý được định dạng ảnh này trên thiết bị."));
    };

    image.src = objectUrl;
  });
}

function drawCompressedBlob(image, width, height, quality) {
  const canvas = document.createElement("canvas");
  const context = canvas.getContext("2d", { alpha: false });
  canvas.width = width;
  canvas.height = height;
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, width, height);
  context.drawImage(image, 0, 0, width, height);

  if (canvas.toBlob) {
    return new Promise((resolve, reject) => {
      canvas.toBlob((blob) => {
        if (blob) {
          resolve(blob);
        } else {
          reject(new Error("Không thể nén ảnh."));
        }
      }, "image/jpeg", quality);
    });
  }

  return Promise.resolve(dataUrlToBlob(canvas.toDataURL("image/jpeg", quality)));
}

function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => {
      const result = String(reader.result || "");
      resolve(result.split(",")[1] || "");
    };
    reader.onerror = () => reject(new Error("Không thể chuyển ảnh để tải lên."));
    reader.readAsDataURL(blob);
  });
}

function dataUrlToBlob(dataUrl) {
  const parts = dataUrl.split(",");
  const mime = parts[0].match(/:(.*?);/)[1];
  const binary = atob(parts[1]);
  const array = new Uint8Array(binary.length);

  for (let index = 0; index < binary.length; index += 1) {
    array[index] = binary.charCodeAt(index);
  }

  return new Blob([array], { type: mime });
}

function buildUploadFileName(originalName) {
  const safeBase = String(originalName || "ao-dai")
    .replace(/\.[^.]+$/, "")
    .replace(/[^\w-]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 40) || "ao-dai";

  return `${safeBase}-${Date.now()}.jpg`;
}

function buildSummaryFromState() {
  return {
    productCount: getCatalogProducts().length,
    totalStock: state.products.reduce((sum, product) => sum + product.currentStock, 0),
    orderCount: state.orders.length
  };
}

function getOrderTotalUnits(order) {
  return (order.items || []).reduce((sum, item) => sum + Number(item.quantity || 0), 0);
}

function getOrderTotalValue(order) {
  return Number(order.totalAmount || 0);
}

function applyUpdatedStock(stock) {
  if (!stock || !stock.sku) {
    return;
  }

  state.products = state.products.map((product) =>
    product.sku === stock.sku
      ? { ...product, currentStock: Number(stock.quantity || 0) }
      : product
  );
}

function normalizeProduct(raw) {
  if (Array.isArray(raw)) {
    const isNewShape = raw.length >= 8 && (!looksLikeImageValue(raw[5]) || Boolean(raw[6]) || Boolean(raw[7]));
    return {
      sku: String(raw[0] || ""),
      name: String(raw[1] || ""),
      material: String(isNewShape ? raw[2] || "" : ""),
      color: String(isNewShape ? raw[3] || "" : raw[2] || ""),
      size: String(isNewShape ? raw[4] || "" : raw[3] || ""),
      price: 0,
      image: String(isNewShape ? raw[6] || "" : raw[5] || ""),
      productId: String(isNewShape ? raw[7] || raw[0] || "" : raw[0] || ""),
      currentStock: Number(isNewShape ? raw[8] || 0 : raw[6] || 0)
    };
  }

  return {
    sku: String(raw.sku || raw.SKU || ""),
    name: String(raw.name || ""),
    material: String(raw.material || ""),
    color: String(raw.color || ""),
    size: String(raw.size || ""),
    price: 0,
    image: String(raw.image || raw.imageUrl || raw.image_url || ""),
    productId: String(raw.productId || raw.product_id || raw.baseSku || raw.sku || raw.SKU || ""),
    currentStock: Number(raw.currentStock || raw.current_stock || raw.stock || 0)
  };
}

function isRenderableProduct(product) {
  return Boolean(product.sku || product.name || product.productId);
}

function looksLikeImageValue(value) {
  return /^https?:\/\//i.test(String(value || "")) || /^data:image\//i.test(String(value || ""));
}

function normalizeOrder(raw) {
  return {
    orderId: String(raw.orderId || raw.order_id || ""),
    date: String(raw.date || raw.createdAt || ""),
    customerName: String(raw.customerName || raw.customer_name || ""),
    customerPhone: String(raw.customerPhone || raw.customer_phone || ""),
    customerAddress: String(raw.customerAddress || raw.customer_address || ""),
    status: normalizeOrderStatus(raw.status),
    totalAmount: Number(raw.totalAmount || raw.total_amount || 0),
    items: Array.isArray(raw.items) ? raw.items.map(normalizeOrderItem) : []
  };
}

function normalizeOrderItem(raw) {
  const product = state.products.find((entry) => entry.sku === String(raw.sku || raw.SKU || ""));
  return {
    sku: String(raw.sku || raw.SKU || ""),
    name: String(raw.name || product?.name || ""),
    color: String(raw.color || product?.color || ""),
    size: String(raw.size || ""),
    quantity: Number(raw.quantity || 0)
  };
}

function normalizeSummary(raw) {
  return {
    productCount: Number(raw.productCount || raw.product_count || getCatalogProducts().length || 0),
    totalStock: Number(
      raw.totalStock ||
      raw.total_stock ||
      state.products.reduce((sum, product) => sum + product.currentStock, 0) ||
      0
    ),
    orderCount: Number(raw.orderCount || raw.order_count || state.orders.length || 0)
  };
}

function getSelectedProduct(sku) {
  return state.products.find((product) => product.sku === sku);
}

function getVisibleProducts() {
  const products = getCatalogProducts();

  if (!state.filter) {
    return products;
  }

  return products.filter((product) => {
    const haystack = [
      product.productId,
      product.name,
      product.material,
      product.colors.join(" "),
      product.sizes.join(" "),
      product.variants.map((item) => item.sku).join(" ")
    ].join(" ").toLowerCase();
    return haystack.includes(state.filter);
  });
}

function getCatalogProducts() {
  const productMap = new Map();

  state.products.forEach((product) => {
    const key = getProductGroupKey(product);
    const existing = productMap.get(key);

    if (!existing) {
      productMap.set(key, {
        productId: key,
        name: product.name,
        material: product.material,
        image: product.image,
        totalStock: product.currentStock,
        variantCount: 1,
        variants: [product],
        colors: product.color ? [product.color] : [],
        sizes: product.size ? [product.size] : []
      });
      return;
    }

    existing.totalStock += product.currentStock;
    existing.variantCount += 1;
    existing.variants.push(product);

    if (product.color && !existing.colors.includes(product.color)) {
      existing.colors.push(product.color);
    }

    if (product.size && !existing.sizes.includes(product.size)) {
      existing.sizes.push(product.size);
    }

    if (!existing.material && product.material) {
      existing.material = product.material;
    }
  });

  return Array.from(productMap.values());
}

function getVariantsByProductId(productId) {
  return state.products.filter((product) => getProductGroupKey(product) === productId);
}

function getProductVariantGroups(product) {
  const groups = new Map();

  (product.variants || []).forEach((variant) => {
    const color = variant.color || "Chưa có màu";
    if (!groups.has(color)) {
      groups.set(color, []);
    }

    groups.get(color).push(variant);
  });

  return Array.from(groups.entries()).map(([color, variants]) => ({
    color,
    variants: variants.slice().sort((left, right) => String(left.size).localeCompare(String(right.size), "vi"))
  }));
}

function getProductGroupKey(product) {
  return String(product.productId || product.sku || "");
}

function renderTagList(items, prefix) {
  const label = `<span class="tag is-label">${escapeHtml(prefix)}</span>`;

  if (!items.length) {
    return `${label}<span class="tag">Chưa có</span>`;
  }

  return label.concat(
    items
      .slice(0, 5)
      .map((item) => `<span class="tag">${escapeHtml(item)}</span>`)
      .concat(items.length > 5 ? [`<span class="tag">+${items.length - 5}</span>`] : [])
      .join("")
  );
}

function getStockClass(quantity) {
  if (quantity <= 0) {
    return "stock-chip out";
  }

  if (quantity <= 3) {
    return "stock-chip low";
  }

  return "stock-chip";
}

function renderEmptyState(title, description) {
  return `
    <div class="empty-state">
      <div>
        <strong>${escapeHtml(title)}</strong>
        <p>${escapeHtml(description)}</p>
      </div>
    </div>
  `;
}

function setStatus(message, type = "info") {
  showToast(message, type);
}

function showToast(message, type = "info") {
  if (!el.toastViewport || !message) {
    return;
  }

  const toastKey = `${type}:${message}`;
  const existingToast = Array.from(el.toastViewport.children).find((toast) => toast.dataset.toastKey === toastKey);
  if (existingToast) {
    scheduleToastDismiss(existingToast, type);
    return;
  }

  while (el.toastViewport.children.length >= 3) {
    const oldestToast = el.toastViewport.firstElementChild;
    if (!oldestToast) {
      break;
    }

    if (oldestToast.dismissTimer) {
      window.clearTimeout(oldestToast.dismissTimer);
    }

    oldestToast.remove();
  }

  const toast = document.createElement("div");
  toast.className = `toast ${type}`.trim();
  toast.dataset.toastKey = toastKey;
  toast.setAttribute("role", type === "error" ? "alert" : "status");
  toast.innerHTML = `
    <div class="toast-message">${escapeHtml(message)}</div>
    <button type="button" class="toast-close" data-toast-close aria-label="Đóng thông báo">&times;</button>
  `;

  el.toastViewport.appendChild(toast);

  requestAnimationFrame(() => {
    toast.classList.add("is-visible");
  });

  scheduleToastDismiss(toast, type);
}

function scheduleToastDismiss(toast, type = "info") {
  if (!toast) {
    return;
  }

  if (toast.dismissTimer) {
    window.clearTimeout(toast.dismissTimer);
  }

  const duration = {
    success: 2400,
    info: 2200,
    error: 3600
  }[type] || 2400;

  toast.dismissTimer = window.setTimeout(() => {
    dismissToast(toast);
  }, duration);
}

function dismissToast(toast) {
  if (!toast || toast.dataset.closing === "true") {
    return;
  }

  toast.dataset.closing = "true";

  if (toast.dismissTimer) {
    window.clearTimeout(toast.dismissTimer);
  }

  toast.classList.remove("is-visible");

  const removeToast = () => {
    if (toast.parentNode) {
      toast.parentNode.removeChild(toast);
    }
  };

  toast.addEventListener("transitionend", removeToast, { once: true });
  window.setTimeout(removeToast, 240);
}

function setBusy(mode, active) {
  state.isBusy = active;
  el.refreshButton.disabled = active;
  el.productSubmitButton.disabled = active;
  el.addVariantButton.disabled = active;
  el.addStockVariantButton.disabled = active;
  el.stockSubmitButton.disabled = active;
  el.addItemButton.disabled = active;
  el.submitOrderButton.disabled = active;

  el.productSubmitButton.textContent = active && mode === "product" ? "Đang lưu..." : "Lưu mẫu áo dài";
  el.stockSubmitButton.textContent = active && mode === "stock" ? "Đang lưu..." : "Lưu nhập kho";
  el.submitOrderButton.textContent = active
    ? (mode === "order" ? "Đang lưu..." : "Đang xử lý...")
    : (state.orderModalMode === "edit" ? "Lưu chỉnh sửa" : "Tạo đơn hàng");
  showLoadingOverlay(active, mode);
}

function setLoadingState(active, mode) {
  el.refreshButton.disabled = active || state.isBusy;
  showLoadingOverlay(active, mode);
  if (!active && state.isBusy) {
    el.refreshButton.disabled = true;
  }
}

function renderProductDraftMeta() {
  if (!el.productDraftCount || !el.productDraftStock) {
    return;
  }

  const totalVariants = state.productVariants.length;
  const totalStock = state.productVariants.reduce((sum, variant) => sum + Number(variant.quantity || 0), 0);

  el.productDraftCount.textContent = totalVariants ? `${totalVariants} biến thể` : "0 biến thể";
  el.productDraftStock.textContent = `Tồn kho nhập ngay: ${totalStock}`;
}

function renderStockDraftMeta() {
  const totalLines = state.stockVariants.length;
  const totalQuantity = state.stockVariants.reduce((sum, variant) => sum + Number(variant.quantity || 0), 0);

  el.stockDraftCount.textContent = totalLines ? `${totalLines} dòng` : "0 dòng";
  el.stockDraftQuantity.textContent = `Tổng số lượng nhập: ${totalQuantity}`;
}

function resetProductForm() {
  state.productVariants = [];
  el.productForm.reset();
  resetVariantComposer();
  clearPendingImage();
  renderImagePreview();
  renderVariantDraft();
  renderVariantPreview();
}

function resetOrderForm() {
  state.orderDraft = [];
  state.selectedOrderSku = "";
  el.orderItemForm.reset();
  el.orderProductSearch.value = "";
  el.orderProductResults.innerHTML = "";
  el.orderQuantity.value = "1";
  el.customerName.value = "";
  el.customerPhone.value = "";
  el.customerAddress.value = "";
  el.orderStatus.value = ORDER_STATUS.PENDING;
  el.orderTotalAmount.value = "";
}

function cloneOrderItem(item) {
  return {
    sku: item.sku,
    name: item.name,
    color: item.color,
    size: item.size,
    quantity: item.quantity
  };
}

function getOriginalOrderQuantityForSku(sku) {
  return state.orderOriginalItems
    .filter((item) => item.sku === sku)
    .reduce((sum, item) => sum + Number(item.quantity || 0), 0);
}

function getAvailableStockForOrderSku(sku) {
  const product = getSelectedProduct(sku);
  if (!product) {
    return 0;
  }

  return Number(product.currentStock || 0) + getOriginalOrderQuantityForSku(sku);
}

function resetStockComposer(options = {}) {
  const keepColor = Boolean(options.keepColor);
  const preferredVariant = options.preferredVariant || null;
  if (!keepColor) {
    el.stockColor.value = preferredVariant ? preferredVariant.color : "";
  }
  el.stockSize.value = preferredVariant ? preferredVariant.size : "";
  el.stockQuantity.value = "1";
  syncStockVariantDefaults();
}

function clearPendingImage() {
  if (state.pendingImage && state.pendingImage.previewUrl) {
    URL.revokeObjectURL(state.pendingImage.previewUrl);
  }

  state.pendingImage = null;
}

function resetVariantComposer(options = {}) {
  const keepColor = Boolean(options.keepColor);
  if (!keepColor) {
    el.variantColor.value = "";
  }
  el.variantSize.value = "";
  el.variantQuantity.value = "0";
}

function getExistingVariantByDraftInput() {
  if (!state.stockTargetProductId) {
    return null;
  }

  const color = el.stockColor.value.trim().toLowerCase();
  const size = el.stockSize.value.trim().toLowerCase();
  if (!color || !size) {
    return null;
  }

  return getVariantsByProductId(state.stockTargetProductId).find(
    (variant) => variant.color.toLowerCase() === color && variant.size.toLowerCase() === size
  ) || null;
}

function upsertProductsState(products) {
  if (!products.length) {
    return;
  }

  const map = new Map(state.products.map((product) => [product.sku, product]));
  products.forEach((product) => {
    map.set(product.sku, product);
  });
  state.products = Array.from(map.values());
}

function formatCurrency(value) {
  return `${money.format(Number(value || 0))} đ`;
}

function formatPriceRange(minValue, maxValue) {
  const min = Number(minValue || 0);
  const max = Number(maxValue || 0);

  if (min === max) {
    return formatCurrency(min);
  }

  return `${formatCurrency(min)} - ${formatCurrency(max)}`;
}

function humanFileSize(bytes) {
  if (bytes < 1024) {
    return `${bytes} B`;
  }

  return `${(bytes / 1024).toFixed(1)} KB`;
}

function formatTime(value) {
  return new Date(value).toLocaleTimeString("vi-VN", {
    hour: "2-digit",
    minute: "2-digit"
  });
}

function formatDateTime(value) {
  if (!value) {
    return "Chưa có thời gian";
  }

  return new Date(value).toLocaleString("vi-VN", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit"
  });
}

function normalizePhone(value) {
  return String(value || "").replace(/[^\d+]/g, "");
}

function normalizeText(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();
}

function normalizeOrderStatus(value) {
  return value === ORDER_STATUS.SENT ? ORDER_STATUS.SENT : ORDER_STATUS.PENDING;
}

function getOrderStatusMeta(status) {
  if (status === ORDER_STATUS.SENT) {
    return {
      label: "Đã giao cho đơn vị vận chuyển",
      shortLabel: "Đã giao",
      className: "sent",
      icon: "✓"
    };
  }

  return {
    label: "Chưa giao cho đơn vị vận chuyển",
    shortLabel: "Chờ giao",
    className: "pending",
    icon: "✓"
  };
}

function showLoadingOverlay(active, mode) {
  const meta = {
    product: {
      title: "Đang lưu mẫu áo dài",
      message: "App đang tải ảnh lên và tạo toàn bộ biến thể màu-size cho mẫu này. Vui lòng chờ trong giây lát."
    },
    stock: {
      title: "Đang lưu nhập kho",
      message: "Các dòng màu, size và số lượng đang được ghi vào mẫu áo dài này."
    },
    order: {
      title: state.orderModalMode === "edit" ? "Đang cập nhật đơn hàng" : "Đang tạo đơn hàng",
      message: "Đơn hàng và tồn kho đang được cập nhật đồng thời."
    },
    "order-delete": {
      title: "Đang xóa đơn hàng",
      message: "Đơn hàng đang được xóa và tồn kho đang được hoàn lại."
    },
    "order-status": {
      title: "Đang cập nhật trạng thái",
      message: "Trạng thái giao cho đơn vị vận chuyển đang được lưu."
    },
    "sync-products": {
      title: "Đang tải dữ liệu",
      message: "App đang lấy danh sách sản phẩm và tồn kho mới nhất."
    },
    "sync-orders": {
      title: "Đang tải đơn hàng",
      message: "App đang lấy danh sách đơn hàng mới nhất."
    },
    "sync-all": {
      title: "Đang đồng bộ dữ liệu",
      message: "App đang lấy lại sản phẩm, tồn kho và đơn hàng mới nhất."
    }
  }[mode] || {
    title: "Đang xử lý",
    message: "Vui lòng chờ trong giây lát..."
  };

  el.loadingTitle.textContent = meta.title;
  el.loadingMessage.textContent = meta.message;
  el.loadingOverlay.hidden = !active;
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

async function disableClientCache() {
  try {
    localStorage.removeItem("ao-dai-inventory-cache-v5");
  } catch (error) {
    console.warn("Không thể xóa dữ liệu local cũ", error);
  }

  if ("caches" in window) {
    try {
      const keys = await caches.keys();
      await Promise.all(
        keys
          .filter((key) => key.startsWith("ao-dai-"))
          .map((key) => caches.delete(key))
      );
    } catch (error) {
      console.warn("Không thể xóa cache trình duyệt cũ", error);
    }
  }

  if (!("serviceWorker" in navigator)) {
    return;
  }

  try {
    const registrations = await navigator.serviceWorker.getRegistrations();
    await Promise.all(registrations.map((registration) => registration.unregister()));
  } catch (error) {
    console.warn("Không thể gỡ service worker cũ", error);
  }
}
