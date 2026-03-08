;(function bootstrapAutoPanel() {
	if (window.__CHECKOWNER_AUTO_PANEL__) {
		window.__CHECKOWNER_AUTO_PANEL__.handleRouteChange()
		return
	}

	const TARGET_HOST = "owner.jiangongdata.com"
	const TARGET_PATH = "/registerPerson"
	const PANEL_ID = "checkowner-auto-panel-host"

	const state = {
		popupState: null,
		selectedQueryIds: new Set(),
		collapsed: false,
		busy: false,
	}

	let host = null
	let shadowRootRef = null
	let refs = {}
	let currentUrl = location.href

	window.__CHECKOWNER_AUTO_PANEL__ = {
		handleRouteChange,
	}

	patchHistory()
	window.addEventListener("popstate", handleRouteChange)
	chrome.storage.onChanged.addListener(handleStorageChange)

	if (document.readyState === "loading") {
		document.addEventListener("DOMContentLoaded", handleRouteChange, { once: true })
	} else {
		handleRouteChange()
	}

	function handleStorageChange(changes, areaName) {
		if (areaName !== "local" || !host) {
			return
		}

		if (changes.queryHistory || changes.authToken || changes.authMeta || changes.exportState) {
			void loadState(false)
		}
	}

	function handleRouteChange() {
		if (currentUrl === location.href && host) {
			return
		}

		currentUrl = location.href

		if (!isTargetPage(location.href)) {
			unmountPanel()
			return
		}

		mountPanel()
		void loadState(true)
	}

	function mountPanel() {
		if (host && shadowRootRef) {
			setCollapsed(false)
			return
		}

		state.collapsed = false
		host = document.createElement("div")
		host.id = PANEL_ID
		host.style.all = "initial"

		shadowRootRef = host.attachShadow({ mode: "open" })
		shadowRootRef.appendChild(buildStyles())
		shadowRootRef.appendChild(buildPanel())

		const mountPoint = document.body || document.documentElement
		mountPoint.appendChild(host)
		setCollapsed(false)

		requestAnimationFrame(() => {
			refs.shell?.classList.add("is-visible")
		})
	}

	function unmountPanel() {
		if (!host) {
			return
		}

		host.remove()
		host = null
		shadowRootRef = null
		refs = {}
	}

	function buildStyles() {
		const style = document.createElement("style")
		style.textContent = `
			:host {
				all: initial;
			}

			* {
				box-sizing: border-box;
			}

			.dock {
				position: fixed;
				inset-block-start: 20px;
				inset-inline-end: 18px;
				z-index: 2147483646;
				font-family: "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", sans-serif;
				color: #2f2618;
				pointer-events: none;
			}

			.shell {
				display: flex;
				align-items: stretch;
				gap: 0;
				transform: translateX(22px);
				opacity: 0;
				transition:
					transform 220ms cubic-bezier(0.22, 1, 0.36, 1),
					opacity 220ms ease;
				will-change: transform, opacity;
				pointer-events: none;
			}

			.shell.is-visible {
				transform: translateX(0);
				opacity: 1;
			}

			.rail {
				inline-size: 54px;
				border: 1px solid rgba(54, 68, 54, 0.14);
				border-inline-end: none;
				border-radius: 18px 0 0 18px;
				background:
					linear-gradient(180deg, rgba(239, 246, 233, 0.96), rgba(223, 235, 227, 0.96));
				box-shadow:
					0 14px 34px rgba(47, 38, 24, 0.18),
					inset 0 1px 0 rgba(255, 255, 255, 0.68);
				display: flex;
				align-items: center;
				justify-content: center;
				cursor: pointer;
				padding: 10px 6px;
				color: #184c43;
				font-weight: 700;
				letter-spacing: 0.1em;
				writing-mode: vertical-rl;
				text-orientation: mixed;
				user-select: none;
				pointer-events: auto;
				transition:
					background-color 160ms ease,
					color 160ms ease,
					inline-size 180ms ease,
					block-size 180ms ease,
					border-radius 180ms ease,
					padding 180ms ease,
					box-shadow 180ms ease,
					border-color 180ms ease;
			}

			.rail:hover {
				background:
					linear-gradient(180deg, rgba(226, 239, 226, 0.98), rgba(211, 229, 221, 0.98));
			}

			.shell.is-collapsed .rail {
				inline-size: 52px;
				block-size: 52px;
				border-inline-end: 1px solid rgba(54, 68, 54, 0.14);
				border-radius: 999px;
				padding: 0;
				writing-mode: horizontal-tb;
				text-orientation: initial;
				letter-spacing: 0;
				font-size: 16px;
				box-shadow:
					0 14px 28px rgba(47, 38, 24, 0.16),
					inset 0 1px 0 rgba(255, 255, 255, 0.7);
			}

			.shell.is-collapsed .panel {
				display: none !important;
				pointer-events: none !important;
			}

			.panel {
				inline-size: 332px;
				max-block-size: calc(100vh - 40px);
				overflow: hidden;
				border: 1px solid rgba(98, 76, 51, 0.16);
				border-radius: 0 24px 24px 24px;
				background:
					radial-gradient(circle at top right, rgba(239, 204, 138, 0.44), transparent 42%),
					linear-gradient(180deg, rgba(255, 251, 243, 0.98), rgba(247, 240, 226, 0.98));
				box-shadow:
					0 18px 42px rgba(47, 38, 24, 0.18),
					inset 0 1px 0 rgba(255, 255, 255, 0.76);
				backdrop-filter: blur(12px);
				display: flex;
				flex-direction: column;
				pointer-events: auto;
			}

			.header {
				padding: 18px 18px 14px;
				border-block-end: 1px solid rgba(98, 76, 51, 0.12);
				background:
					linear-gradient(180deg, rgba(255, 255, 255, 0.56), rgba(255, 255, 255, 0));
			}

			.headline {
				display: flex;
				align-items: flex-start;
				justify-content: space-between;
				gap: 12px;
			}

			.kicker {
				margin: 0 0 6px;
				font-size: 11px;
				line-height: 1.3;
				letter-spacing: 0.16em;
				text-transform: uppercase;
				color: #6e604c;
			}

			.title {
				margin: 0;
				font-family: "Songti SC", "STSong", "Noto Serif SC", serif;
				font-size: 22px;
				line-height: 1.15;
				letter-spacing: -0.02em;
				color: #2e261a;
			}

			.body {
				padding: 16px 16px 18px;
				display: grid;
				gap: 14px;
				overflow: auto;
			}

			.section {
				display: grid;
				gap: 10px;
			}

			.section-head {
				display: flex;
				align-items: center;
				justify-content: space-between;
				gap: 12px;
			}

			.section-title {
				margin: 0;
				font-size: 15px;
				font-weight: 800;
				color: #2e261a;
			}

			.meta {
				margin: 0;
				font-size: 12px;
				line-height: 1.5;
				color: #72634d;
			}

			.query-list {
				display: grid;
				gap: 10px;
				max-block-size: 310px;
				overflow: auto;
				padding-inline-end: 2px;
			}

			.query-item {
				display: flex;
				gap: 10px;
				align-items: flex-start;
				padding: 12px;
				border-radius: 16px;
				border: 1px solid rgba(98, 76, 51, 0.12);
				background: rgba(255, 253, 248, 0.9);
				transition:
					transform 140ms ease,
					border-color 140ms ease,
					background-color 140ms ease;
			}

			.query-item:hover {
				transform: translateY(-1px);
				border-color: rgba(15, 109, 91, 0.22);
			}

			.query-item.is-selected {
				background: rgba(215, 239, 232, 0.75);
				border-color: rgba(15, 109, 91, 0.28);
			}

			.query-checkbox {
				margin: 3px 0 0;
				accent-color: #0f6d5b;
				inline-size: 16px;
				block-size: 16px;
				flex: 0 0 auto;
			}

			.query-content {
				min-inline-size: 0;
			}

			.query-name {
				margin: 0;
				font-size: 14px;
				line-height: 1.45;
				font-weight: 700;
				color: #2f2618;
			}

			.empty {
				padding: 16px 14px;
				border-radius: 16px;
				border: 1px dashed rgba(98, 76, 51, 0.24);
				background: rgba(255, 255, 255, 0.48);
				font-size: 12px;
				line-height: 1.7;
				color: #7a6b56;
			}

			.actions,
			.utility {
				display: flex;
				align-items: center;
				justify-content: space-between;
				gap: 10px;
				flex-wrap: wrap;
			}

			.button-group {
				display: flex;
				gap: 8px;
				flex-wrap: wrap;
			}

			.query-tools {
				display: flex;
				gap: 8px;
				flex-wrap: wrap;
			}

			button {
				appearance: none;
				border: 1px solid transparent;
				border-radius: 999px;
				padding: 10px 14px;
				font-size: 12px;
				line-height: 1;
				font-weight: 800;
				cursor: pointer;
				transition:
					transform 120ms ease,
					border-color 120ms ease,
					background-color 120ms ease,
					color 120ms ease,
					opacity 120ms ease;
			}

			button:hover {
				transform: translateY(-1px);
			}

			button:disabled {
				opacity: 0.45;
				cursor: not-allowed;
				transform: none;
			}

			.button-primary {
				background: #0f6d5b;
				color: #fffdf7;
				box-shadow: 0 10px 20px rgba(15, 109, 91, 0.16);
			}

			.button-secondary {
				background: rgba(255, 255, 255, 0.78);
				border-color: rgba(98, 76, 51, 0.16);
				color: #2e261a;
			}

			.button-danger {
				background: rgba(252, 233, 228, 0.92);
				color: #8d2b2b;
			}

			.button-icon {
				inline-size: 34px;
				block-size: 34px;
				padding: 0;
				display: inline-flex;
				align-items: center;
				justify-content: center;
				background: rgba(255, 255, 255, 0.72);
				border-color: rgba(98, 76, 51, 0.16);
				color: #584a38;
				font-size: 16px;
			}

			.status-box {
				padding: 12px 14px;
				border-radius: 16px;
				background: rgba(255, 255, 255, 0.7);
				border: 1px solid rgba(98, 76, 51, 0.1);
			}

			.status-title {
				margin: 0;
				font-size: 14px;
				font-weight: 800;
				line-height: 1.45;
				color: #2f2618;
			}

			.status-detail {
				margin: 6px 0 0;
				font-size: 12px;
				line-height: 1.6;
				color: #736451;
			}

			.progress-track {
				margin: 10px 0 6px;
				inline-size: 100%;
				block-size: 10px;
				border-radius: 999px;
				background: rgba(15, 109, 91, 0.12);
				overflow: hidden;
			}

			.progress-fill {
				inline-size: 0%;
				block-size: 100%;
				border-radius: inherit;
				background: linear-gradient(90deg, #0f6d5b, #61a78d);
				transition: width 180ms ease;
			}

			.progress-text {
				margin: 0;
				font-size: 12px;
				line-height: 1.4;
				font-weight: 700;
				color: #0f6d5b;
			}

			.status-detail.is-error {
				color: #9a3030;
			}

			@media (prefers-reduced-motion: reduce) {
				.shell,
				.query-item,
				button {
					transition: none !important;
				}
			}
		`
		return style
	}

	function buildPanel() {
		const dock = document.createElement("section")
		dock.className = "dock"

		const shell = document.createElement("div")
		shell.className = "shell"
		refs.shell = shell

		const rail = document.createElement("button")
		rail.className = "rail"
		rail.type = "button"
		rail.textContent = "查参助手"
		rail.title = "收起面板"
		rail.setAttribute("aria-label", "收起面板")
		rail.addEventListener("click", () => {
			setCollapsed(!state.collapsed)
		})
		refs.rail = rail

		const panel = document.createElement("section")
		panel.className = "panel"
		refs.panel = panel

		const header = document.createElement("header")
		header.className = "header"

		const headline = document.createElement("div")
		headline.className = "headline"

		const titleWrap = document.createElement("div")
		const kicker = document.createElement("p")
		kicker.className = "kicker"
		kicker.textContent = "人员查询页"

		const title = document.createElement("h1")
		title.className = "title"
		title.textContent = "查业主-Excel导出"

		titleWrap.appendChild(kicker)
		titleWrap.appendChild(title)

		const headActions = document.createElement("div")
		headActions.className = "button-group"

		const collapseButton = document.createElement("button")
		collapseButton.type = "button"
		collapseButton.className = "button-icon"
		collapseButton.textContent = "−"
		collapseButton.title = "收起"
		collapseButton.addEventListener("click", () => {
			setCollapsed(true)
		})
		refs.collapseButton = collapseButton

		headActions.appendChild(collapseButton)
		headline.appendChild(titleWrap)
		headline.appendChild(headActions)

		header.appendChild(headline)

		const body = document.createElement("div")
		body.className = "body"

		const utilitySection = document.createElement("section")
		utilitySection.className = "section"
		const utilityHead = document.createElement("div")
		utilityHead.className = "utility"

		const utilityTextWrap = document.createElement("div")
		const utilityTitle = document.createElement("p")
		utilityTitle.className = "section-title"
		utilityTitle.textContent = "页面状态"
		const utilityMeta = document.createElement("p")
		utilityMeta.className = "meta"
		utilityMeta.textContent = "进入这个页面后面板会自动出现"
		refs.utilityMeta = utilityMeta
		utilityTextWrap.appendChild(utilityTitle)
		utilityTextWrap.appendChild(utilityMeta)

		const utilityButtons = document.createElement("div")
		utilityButtons.className = "button-group"
		const refreshButton = createButton("button-secondary", "刷新")
		refreshButton.addEventListener("click", () => {
			void loadState(true)
		})
		const rescanButton = createButton("button-secondary", "重新扫描")
		rescanButton.addEventListener("click", async () => {
			await runBusyAction(async () => {
				const response = await chrome.runtime.sendMessage({ type: "rescan-active-tab" })
				if (!response?.ok) {
					throw new Error(response?.error || "重新扫描失败")
				}
				await loadState(false)
			}, "重新扫描失败")
		})
		refs.refreshButton = refreshButton
		refs.rescanButton = rescanButton
		utilityButtons.appendChild(refreshButton)
		utilityButtons.appendChild(rescanButton)

		utilityHead.appendChild(utilityTextWrap)
		utilityHead.appendChild(utilityButtons)
		utilitySection.appendChild(utilityHead)

		const querySection = document.createElement("section")
		querySection.className = "section"
		const queryHead = document.createElement("div")
		queryHead.className = "section-head"
		const queryHeadTitle = document.createElement("p")
		queryHeadTitle.className = "section-title"
		queryHeadTitle.textContent = "条件列表"
		const queryTools = document.createElement("div")
		queryTools.className = "query-tools"
		const toggleSelectButton = createButton("button-secondary", "全选")
		toggleSelectButton.addEventListener("click", () => {
			toggleSelectQueries()
		})
		refs.toggleSelectButton = toggleSelectButton
		queryHead.appendChild(queryHeadTitle)
		queryHead.appendChild(queryTools)
		queryTools.appendChild(toggleSelectButton)

		const queryList = document.createElement("div")
		queryList.className = "query-list"
		refs.queryList = queryList
		querySection.appendChild(queryHead)
		querySection.appendChild(queryList)

		const statusSection = document.createElement("section")
		statusSection.className = "section"
		const statusBox = document.createElement("div")
		statusBox.className = "status-box"
		const statusTitle = document.createElement("p")
		statusTitle.className = "status-title"
		statusTitle.textContent = "等待导出"
		const statusDetail = document.createElement("p")
		statusDetail.className = "status-detail"
		statusDetail.textContent = "先在页面执行一次查询，再导出。"
		const progressTrack = document.createElement("div")
		progressTrack.className = "progress-track"
		const progressFill = document.createElement("div")
		progressFill.className = "progress-fill"
		progressTrack.appendChild(progressFill)
		const progressText = document.createElement("p")
		progressText.className = "progress-text"
		progressText.textContent = "0%"
		refs.statusTitle = statusTitle
		refs.statusDetail = statusDetail
		refs.progressFill = progressFill
		refs.progressText = progressText
		statusBox.appendChild(statusTitle)
		statusBox.appendChild(progressTrack)
		statusBox.appendChild(progressText)
		statusBox.appendChild(statusDetail)
		statusSection.appendChild(statusBox)

		const actions = document.createElement("div")
		actions.className = "actions"
		const actionsMeta = document.createElement("p")
		actionsMeta.className = "meta"
		actionsMeta.textContent = "导出文件会进入浏览器默认下载目录"
		const actionButtons = document.createElement("div")
		actionButtons.className = "button-group"

		const clearButton = createButton("button-danger", "清空条件")
		clearButton.addEventListener("click", async () => {
			await runBusyAction(async () => {
				const response = await chrome.runtime.sendMessage({ type: "clear-captured-queries" })
				if (!response?.ok) {
					throw new Error(response?.error || "清空失败")
				}
				state.selectedQueryIds = new Set()
				await loadState(true)
			}, "清空失败")
		})

		const exportButton = createButton("button-primary", "导出选中")
		exportButton.addEventListener("click", async () => {
			const queryIds = Array.from(state.selectedQueryIds)
			if (!queryIds.length) {
				renderInlineStatus("没有选中的条件", "请至少勾选一个条件", true)
				return
			}

			await runBusyAction(async () => {
				const response = await chrome.runtime.sendMessage({
					type: "export-queries",
					queryIds,
				})
				if (!response?.ok) {
					throw new Error(response?.error || "导出失败")
				}
				await loadState(false)
			}, "导出失败")
		})

		refs.clearButton = clearButton
		refs.exportButton = exportButton
		actionButtons.appendChild(clearButton)
		actionButtons.appendChild(exportButton)
		actions.appendChild(actionsMeta)
		actions.appendChild(actionButtons)

		body.appendChild(utilitySection)
		body.appendChild(querySection)
		body.appendChild(statusSection)
		body.appendChild(actions)

		panel.appendChild(header)
		panel.appendChild(body)
		shell.appendChild(rail)
		shell.appendChild(panel)
		dock.appendChild(shell)

		return dock
	}

	function createButton(className, text) {
		const button = document.createElement("button")
		button.type = "button"
		button.className = className
		button.textContent = text
		return button
	}

	async function loadState(resetSelection) {
		if (!host) {
			return
		}

		try {
			const response = await chrome.runtime.sendMessage({ type: "get-popup-state" })
			if (!response?.ok) {
				throw new Error(response?.error || "读取状态失败")
			}
			state.popupState = response.state
			renderState(Boolean(resetSelection))
		} catch (error) {
			renderInlineStatus("读取状态失败", error.message, true)
		}
	}

	function renderState(resetSelection) {
		const popupState = state.popupState || {}
		const queryHistory = Array.isArray(popupState.queryHistory) ? popupState.queryHistory : []

		if (resetSelection) {
			state.selectedQueryIds = new Set()
		} else {
			const nextSelection = new Set()
			queryHistory.forEach((item) => {
				if (state.selectedQueryIds.has(item.id)) {
					nextSelection.add(item.id)
				}
			})
			state.selectedQueryIds = nextSelection
		}

		refs.utilityMeta.textContent = queryHistory.length
			? "已捕获条件，可直接勾选导出"
			: "请先登录页面并执行一次查询"

		renderQueryList(queryHistory)
		renderExportState(popupState.exportState || {})
		setButtonsDisabled(Boolean(popupState.exportState?.running) || state.busy)
	}

	function renderQueryList(queryHistory) {
		refs.queryList.replaceChildren()

		if (!queryHistory.length) {
			const empty = document.createElement("div")
			empty.className = "empty"
			empty.textContent = "还没有捕获到查询条件。先在页面里执行一次检索，再点“重新扫描”或等待页面请求被自动记录。"
			refs.queryList.appendChild(empty)
			updateToggleSelectButton(queryHistory)
			return
		}

		const fragment = document.createDocumentFragment()

		queryHistory.forEach((item) => {
			const label = document.createElement("label")
			label.className = "query-item"
			if (state.selectedQueryIds.has(item.id)) {
				label.classList.add("is-selected")
			}

			const checkbox = document.createElement("input")
			checkbox.className = "query-checkbox"
			checkbox.type = "checkbox"
			checkbox.checked = state.selectedQueryIds.has(item.id)
			checkbox.addEventListener("change", () => {
				if (checkbox.checked) {
					state.selectedQueryIds.add(item.id)
					label.classList.add("is-selected")
				} else {
					state.selectedQueryIds.delete(item.id)
					label.classList.remove("is-selected")
				}
				updateToggleSelectButton(queryHistory)
			})

			const content = document.createElement("div")
			content.className = "query-content"

			const name = document.createElement("p")
			name.className = "query-name"
			name.textContent = item.name || "未命名条件"

			content.appendChild(name)
			label.appendChild(checkbox)
			label.appendChild(content)
			fragment.appendChild(label)
		})

		refs.queryList.appendChild(fragment)
		updateToggleSelectButton(queryHistory)
	}

	function renderExportState(exportState) {
		const detailParts = []
		refs.statusTitle.textContent = exportState.message || "等待导出"
		const progressPercent = Number.isFinite(exportState.progressPercent) ? exportState.progressPercent : 0
		const progressCurrent = Number(exportState.progressCurrent) || 0
		const progressTarget = Number(exportState.progressTarget) || 0
		if (refs.progressFill) {
			refs.progressFill.style.width = `${Math.max(0, Math.min(100, progressPercent))}%`
		}
		if (refs.progressText) {
			refs.progressText.textContent = progressTarget
				? `${progressPercent}% (${progressCurrent}/${progressTarget})`
				: `${progressPercent}%`
		}

		if (typeof exportState.completed === "number" && typeof exportState.total === "number" && exportState.total > 0) {
			detailParts.push(`进度 ${exportState.completed}/${exportState.total}`)
		}

		if (exportState.currentName) {
			detailParts.push(exportState.currentName)
		}

		if (exportState.updatedAt) {
			detailParts.push(formatTime(exportState.updatedAt))
		}

		if (exportState.error) {
			detailParts.push(`错误: ${exportState.error}`)
			refs.statusDetail.classList.add("is-error")
		} else {
			refs.statusDetail.classList.remove("is-error")
		}

		refs.statusDetail.textContent = detailParts.join(" | ") || "选中条件后即可导出"
	}

	function renderInlineStatus(title, detail, isError) {
		if (!refs.statusTitle || !refs.statusDetail) {
			return
		}

		refs.statusTitle.textContent = title
		refs.statusDetail.textContent = detail || ""
		refs.statusDetail.classList.toggle("is-error", Boolean(isError))
	}

	async function runBusyAction(action, fallbackMessage) {
		state.busy = true
		setButtonsDisabled(true)
		try {
			await action()
		} catch (error) {
			renderInlineStatus(fallbackMessage, error.message, true)
		} finally {
			state.busy = false
			const running = Boolean(state.popupState?.exportState?.running)
			setButtonsDisabled(running)
		}
	}

	function setButtonsDisabled(disabled) {
		if (refs.refreshButton) {
			refs.refreshButton.disabled = disabled
		}
		if (refs.rescanButton) {
			refs.rescanButton.disabled = disabled
		}
		if (refs.clearButton) {
			refs.clearButton.disabled = disabled
		}
		if (refs.exportButton) {
			refs.exportButton.disabled = disabled
		}
		if (refs.toggleSelectButton) {
			refs.toggleSelectButton.disabled = disabled || !(state.popupState?.queryHistory || []).length
		}
		if (refs.collapseButton) {
			refs.collapseButton.disabled = false
		}
	}

	function selectAllQueries() {
		const queryHistory = state.popupState?.queryHistory || []
		state.selectedQueryIds = new Set(queryHistory.map((item) => item.id))
		renderQueryList(queryHistory)
	}

	function deselectAllQueries() {
		state.selectedQueryIds = new Set()
		renderQueryList(state.popupState?.queryHistory || [])
	}

	function toggleSelectQueries() {
		const queryHistory = state.popupState?.queryHistory || []
		if (areAllQueriesSelected(queryHistory)) {
			deselectAllQueries()
			return
		}
		selectAllQueries()
	}

	function areAllQueriesSelected(queryHistory) {
		return Boolean(queryHistory.length) && queryHistory.every((item) => state.selectedQueryIds.has(item.id))
	}

	function updateToggleSelectButton(queryHistory) {
		if (!refs.toggleSelectButton) {
			return
		}

		const allSelected = areAllQueriesSelected(queryHistory)
		refs.toggleSelectButton.textContent = allSelected ? "取消全选" : "全选"
		refs.toggleSelectButton.disabled = !(queryHistory || []).length
	}

	function setCollapsed(collapsed) {
		state.collapsed = collapsed
		if (!refs.shell || !refs.collapseButton || !refs.rail || !refs.panel) {
			return
		}

		refs.shell.classList.toggle("is-collapsed", collapsed)
		refs.panel.hidden = collapsed
		refs.panel.style.display = collapsed ? "none" : "flex"
		refs.panel.setAttribute("aria-hidden", collapsed ? "true" : "false")
		refs.rail.textContent = collapsed ? "查" : "查参助手"
		refs.rail.title = collapsed ? "展开面板" : "收起面板"
		refs.rail.setAttribute("aria-label", collapsed ? "展开面板" : "收起面板")
		refs.collapseButton.textContent = collapsed ? "+" : "−"
		refs.collapseButton.title = collapsed ? "展开" : "收起"
	}

	function patchHistory() {
		const originalPushState = history.pushState
		const originalReplaceState = history.replaceState

		history.pushState = function patchedPushState() {
			const result = originalPushState.apply(this, arguments)
			queueMicrotask(handleRouteChange)
			return result
		}

		history.replaceState = function patchedReplaceState() {
			const result = originalReplaceState.apply(this, arguments)
			queueMicrotask(handleRouteChange)
			return result
		}
	}

	function isTargetPage(url) {
		try {
			const parsed = new URL(url)
			if (parsed.hostname !== TARGET_HOST) {
				return false
			}

			const pathname = normalizePath(parsed.pathname)
			if (pathname === TARGET_PATH) {
				return true
			}

			return normalizeHashPath(parsed.hash) === TARGET_PATH
		} catch (_error) {
			return false
		}
	}

	function normalizePath(pathname) {
		if (!pathname || pathname === "/") {
			return pathname || "/"
		}

		return pathname.replace(/\/+$/, "")
	}

	function normalizeHashPath(hash) {
		if (!hash) {
			return ""
		}

		const rawHash = String(hash).replace(/^#/, "")
		if (!rawHash) {
			return ""
		}

		const hashPath = rawHash.startsWith("/") ? rawHash : `/${rawHash.replace(/^!?\//, "")}`
		return normalizePath(hashPath.split(/[?#]/, 1)[0])
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
})()
