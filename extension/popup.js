const elements = {
	pageStatus: document.getElementById("pageStatus"),
	pageHint: document.getElementById("pageHint"),
	queryList: document.getElementById("queryList"),
	exportMessage: document.getElementById("exportMessage"),
	exportProgressBar: document.getElementById("exportProgressBar"),
	exportProgressText: document.getElementById("exportProgressText"),
	exportDetail: document.getElementById("exportDetail"),
	refreshButton: document.getElementById("refreshButton"),
	rescanButton: document.getElementById("rescanButton"),
	toggleSelectButton: document.getElementById("toggleSelectButton"),
	clearButton: document.getElementById("clearButton"),
	exportButton: document.getElementById("exportButton"),
}

let popupState = null
let selectedQueryIds = new Set()

init()

function init() {
	bindEvents()
	void loadState()

	chrome.storage.onChanged.addListener((changes, areaName) => {
		if (areaName !== "local") {
			return
		}

		if (changes.queryHistory || changes.authToken || changes.authMeta || changes.exportState) {
			void loadState(false)
		}
	})
}

function bindEvents() {
	elements.refreshButton.addEventListener("click", () => {
		void loadState()
	})

	elements.rescanButton.addEventListener("click", async () => {
		setButtonsDisabled(true)
		try {
			const response = await chrome.runtime.sendMessage({ type: "rescan-active-tab" })
			if (!response?.ok) {
				throw new Error(response?.error || "重新扫描失败")
			}
			await loadState()
		} catch (error) {
			renderExportState({
				message: "重新扫描失败",
				error: error.message,
			})
		} finally {
			setButtonsDisabled(false)
		}
	})

	elements.toggleSelectButton.addEventListener("click", () => {
		toggleSelectQueries()
	})

	elements.clearButton.addEventListener("click", async () => {
		setButtonsDisabled(true)
		try {
			const response = await chrome.runtime.sendMessage({ type: "clear-captured-queries" })
			if (!response?.ok) {
				throw new Error(response?.error || "清空失败")
			}
			selectedQueryIds = new Set()
			await loadState()
		} catch (error) {
			renderExportState({
				message: "清空失败",
				error: error.message,
			})
		} finally {
			setButtonsDisabled(false)
		}
	})

	elements.exportButton.addEventListener("click", async () => {
		const queryIds = Array.from(getSelectedQueryIds())
		if (!queryIds.length) {
			renderExportState({
				message: "没有选中的条件",
				error: "请至少勾选一个条件",
			})
			return
		}

		setButtonsDisabled(true)
		try {
			const response = await chrome.runtime.sendMessage({
				type: "export-queries",
				queryIds,
			})

			if (!response?.ok) {
				throw new Error(response?.error || "导出失败")
			}

			await loadState()
		} catch (error) {
			renderExportState({
				message: "导出失败",
				error: error.message,
			})
		} finally {
			setButtonsDisabled(false)
		}
	})
}

async function loadState(resetSelection = true) {
	const response = await chrome.runtime.sendMessage({ type: "get-popup-state" })
	if (!response?.ok) {
		throw new Error(response?.error || "读取扩展状态失败")
	}

	popupState = response.state
	renderState(resetSelection)
}

function renderState(resetSelection) {
	renderPageStatus()
	renderQueries(resetSelection)
	renderExportState(popupState.exportState)
}

function renderPageStatus() {
	if (!popupState.activeTabUrl) {
		elements.pageStatus.textContent = "未找到当前标签页"
		elements.pageHint.textContent = "请先打开人员查询页面"
		elements.rescanButton.disabled = true
		return
	}

	elements.pageStatus.textContent = popupState.activeTabSupported
		? popupState.activeTabTitle || "人员查询页面"
		: "当前页不是人员查询页面"
	elements.rescanButton.disabled = !popupState.activeTabSupported
	elements.pageHint.textContent = popupState.queryHistory?.length
		? "已捕获条件，可直接勾选导出"
		: "请先登录页面并执行一次查询"
}

function renderQueries(resetSelection) {
	const queryHistory = popupState.queryHistory || []

	if (resetSelection) {
		selectedQueryIds = new Set()
	} else {
		const nextSelection = new Set()
		queryHistory.forEach((item) => {
			if (selectedQueryIds.has(item.id)) {
				nextSelection.add(item.id)
			}
		})
		selectedQueryIds = nextSelection
	}

	if (!queryHistory.length) {
		elements.queryList.innerHTML =
			'<div class="empty-state">还没有捕获到查询条件。先在页面里执行一次检索，再点“重新扫描”或“刷新”。</div>'
		updateToggleSelectButton(queryHistory)
		return
	}

	elements.queryList.innerHTML = ""
	const fragment = document.createDocumentFragment()

	queryHistory.forEach((item) => {
		const wrapper = document.createElement("label")
		wrapper.className = "query-item"
		if (selectedQueryIds.has(item.id)) {
			wrapper.classList.add("selected")
		}

		const queryMain = document.createElement("div")
		queryMain.className = "query-main"

		const checkbox = document.createElement("input")
		checkbox.type = "checkbox"
		checkbox.checked = selectedQueryIds.has(item.id)
		checkbox.addEventListener("change", () => {
			if (checkbox.checked) {
				selectedQueryIds.add(item.id)
				wrapper.classList.add("selected")
			} else {
				selectedQueryIds.delete(item.id)
				wrapper.classList.remove("selected")
			}
			updateToggleSelectButton(queryHistory)
		})

		const content = document.createElement("div")
		const title = document.createElement("p")
		title.className = "query-name"
		title.textContent = item.name || "未命名条件"

		content.appendChild(title)
		queryMain.appendChild(checkbox)
		queryMain.appendChild(content)
		wrapper.appendChild(queryMain)
		fragment.appendChild(wrapper)
	})

	elements.queryList.appendChild(fragment)
	updateToggleSelectButton(queryHistory)
}

function renderExportState(exportState) {
	const state = exportState || {}
	elements.exportMessage.textContent = state.message || "等待导出"
	const progressPercent = Number.isFinite(state.progressPercent) ? state.progressPercent : 0
	const progressCurrent = Number(state.progressCurrent) || 0
	const progressTarget = Number(state.progressTarget) || 0
	elements.exportProgressBar.style.width = `${Math.max(0, Math.min(100, progressPercent))}%`
	elements.exportProgressText.textContent = progressTarget
		? `${progressPercent}% (${progressCurrent}/${progressTarget})`
		: `${progressPercent}%`

	const detailParts = []
	if (typeof state.completed === "number" && typeof state.total === "number" && state.total > 0) {
		detailParts.push(`进度 ${state.completed}/${state.total}`)
	}
	if (state.currentName) {
		detailParts.push(state.currentName)
	}
	if (state.updatedAt) {
		detailParts.push(formatTime(state.updatedAt))
	}
	if (state.error) {
		detailParts.push(`错误: ${state.error}`)
		elements.exportDetail.classList.add("error-text")
	} else {
		elements.exportDetail.classList.remove("error-text")
	}

	elements.exportDetail.textContent = detailParts.join(" | ") || "选中条件后即可导出"
	const disabled = Boolean(state.running)
	setButtonsDisabled(disabled, false)
}

function getSelectedQueryIds() {
	return new Set(selectedQueryIds)
}

function selectAllQueries() {
	const queryHistory = popupState?.queryHistory || []
	selectedQueryIds = new Set(queryHistory.map((item) => item.id))
	renderQueries(false)
}

function deselectAllQueries() {
	selectedQueryIds = new Set()
	renderQueries(false)
}

function toggleSelectQueries() {
	const queryHistory = popupState?.queryHistory || []
	if (areAllQueriesSelected(queryHistory)) {
		deselectAllQueries()
		return
	}
	selectAllQueries()
}

function areAllQueriesSelected(queryHistory) {
	return Boolean(queryHistory.length) && queryHistory.every((item) => selectedQueryIds.has(item.id))
}

function updateToggleSelectButton(queryHistory) {
	if (!elements.toggleSelectButton) {
		return
	}

	const allSelected = areAllQueriesSelected(queryHistory)
	elements.toggleSelectButton.textContent = allSelected ? "取消全选" : "全选"
	elements.toggleSelectButton.disabled = !queryHistory.length
}

function setButtonsDisabled(disabled, includeRefresh = true) {
	elements.exportButton.disabled = disabled
	elements.clearButton.disabled = disabled
	elements.toggleSelectButton.disabled = disabled || !(popupState?.queryHistory || []).length
	if (includeRefresh) {
		elements.refreshButton.disabled = disabled
		elements.rescanButton.disabled = disabled || !popupState?.activeTabSupported
	}
}

function formatTime(value) {
	if (!value) {
		return "未知时间"
	}

	try {
		return new Intl.DateTimeFormat("zh-CN", {
			month: "2-digit",
			day: "2-digit",
			hour: "2-digit",
			minute: "2-digit",
		}).format(new Date(value))
	} catch (_error) {
		return String(value)
	}
}
