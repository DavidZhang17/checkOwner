/* global XLSX, importScripts */
importScripts("lib/xlsx.js")

const API_BASE_URL = "https://gwowner.jiangongdata.com"
const PERSONNEL_ENDPOINT = "/jian-butler-owner-biz/personnel/pagePersonnel"
const PROJECT_ENDPOINT = "/jian-butler-owner-biz/person/achievement/pageProjectWinningByPerson"
const QUERY_HISTORY_LIMIT = 30
const MAX_RECORDS = 5000
const EXPORT_PAGE_SIZE = 5000
const RECOVERED_EXPORT_MESSAGE = "上次导出任务已中断，请重新导出"

const STORAGE_KEYS = {
	AUTH_TOKEN: "authToken",
	AUTH_META: "authMeta",
	QUERY_HISTORY: "queryHistory",
	EXPORT_STATE: "exportState",
}

let activeExportSessionId = ""

chrome.runtime.onInstalled.addListener(async () => {
	const currentState = await chrome.storage.local.get([STORAGE_KEYS.QUERY_HISTORY, STORAGE_KEYS.EXPORT_STATE])
	const exportState = await recoverExportStateIfNeeded(currentState[STORAGE_KEYS.EXPORT_STATE])
	await chrome.storage.local.set({
		[STORAGE_KEYS.QUERY_HISTORY]: Array.isArray(currentState[STORAGE_KEYS.QUERY_HISTORY])
			? currentState[STORAGE_KEYS.QUERY_HISTORY]
			: [],
		[STORAGE_KEYS.EXPORT_STATE]: exportState,
	})
	await harvestCookiesForTokens()
})

chrome.runtime.onStartup.addListener(() => {
	void recoverExportStateIfNeeded()
})

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
	if (!message?.type) {
		return false
	}

	if (message.type === "bridge-event") {
		handleBridgeEvent(message.eventType, message.payload, sender)
			.then(() => sendResponse({ ok: true }))
			.catch((error) => sendResponse({ ok: false, error: error.message }))
		return true
	}

	if (message.type === "get-popup-state") {
		getPopupState()
			.then((state) => sendResponse({ ok: true, state }))
			.catch((error) => sendResponse({ ok: false, error: error.message }))
		return true
	}

	if (message.type === "rescan-active-tab") {
		rescanActiveTab()
			.then((result) => sendResponse({ ok: true, result }))
			.catch((error) => sendResponse({ ok: false, error: error.message }))
		return true
	}

	if (message.type === "clear-captured-queries") {
		chrome.storage.local
			.set({ [STORAGE_KEYS.QUERY_HISTORY]: [] })
			.then(() => sendResponse({ ok: true }))
			.catch((error) => sendResponse({ ok: false, error: error.message }))
		return true
	}

	if (message.type === "export-queries") {
		exportQueries(message.queryIds || [])
			.then((result) => sendResponse({ ok: true, result }))
			.catch((error) => sendResponse({ ok: false, error: error.message }))
		return true
	}

	return false
})

if (chrome.webRequest?.onBeforeSendHeaders) {
	chrome.webRequest.onBeforeSendHeaders.addListener(
		(details) => {
			const authorizationHeader = (details.requestHeaders || []).find(
				(header) => header && header.name && header.name.toLowerCase() === "authorization"
			)

			if (!authorizationHeader?.value) {
				return
			}

			void persistAuthToken(authorizationHeader.value, {
				source: "webRequest",
				url: details.url,
			})
		},
		{
			urls: ["*://*.jiangongdata.com/*"],
		},
		["requestHeaders", "extraHeaders"]
	)
}

async function handleBridgeEvent(eventType, payload = {}, sender) {
	if (!eventType) {
		return
	}

	if (eventType === "page-scan") {
		await ingestTokenCandidates(payload.tokenCandidates || [], payload)
		return
	}

	if (eventType === "auth-captured") {
		await persistAuthToken(payload.authorization, {
			source: payload.source || "page-bridge",
			url: payload.url || sender?.tab?.url || "",
		})
		return
	}

	if (eventType === "query-captured") {
		if (payload.hasAuthorization && payload.authorization) {
			await persistAuthToken(payload.authorization, {
				source: payload.source || "page-bridge",
				url: payload.url || sender?.tab?.url || "",
			})
		}
		return
	}

	if (eventType === "template-list-captured") {
		await replaceQueryHistoryFromTemplates(
			Array.isArray(payload.templates) ? payload.templates : [],
			payload,
			sender
		)
	}
}

async function getPopupState() {
	await harvestCookiesForTokens()

	const storage = await chrome.storage.local.get(Object.values(STORAGE_KEYS))
	const [activeTab] = await chrome.tabs.query({ active: true, currentWindow: true })
	const exportState = await recoverExportStateIfNeeded(storage[STORAGE_KEYS.EXPORT_STATE])

	return {
		authToken: storage[STORAGE_KEYS.AUTH_TOKEN] || "",
		authMeta: storage[STORAGE_KEYS.AUTH_META] || null,
		queryHistory: sortQueries(storage[STORAGE_KEYS.QUERY_HISTORY] || []),
		exportState,
		activeTabUrl: activeTab?.url || "",
		activeTabTitle: activeTab?.title || "",
		activeTabSupported: isSupportedPage(activeTab?.url || ""),
	}
}

async function rescanActiveTab() {
	const [activeTab] = await chrome.tabs.query({ active: true, currentWindow: true })
	if (!activeTab?.id || !isSupportedPage(activeTab.url || "")) {
		throw new Error("请先打开建工数据相关页面，再点击重新扫描")
	}

	await chrome.scripting.executeScript({
		target: { tabId: activeTab.id },
		files: ["content-script.js"],
	})

	try {
		await chrome.tabs.sendMessage(activeTab.id, { type: "rescanPage" })
	} catch (_error) {
		// executeScript has already injected the content script; a missing listener can be ignored
	}

	await harvestCookiesForTokens()

	return { rescanned: true }
}

async function harvestCookiesForTokens() {
	const cookies = await chrome.cookies.getAll({ domain: "jiangongdata.com" })
	const candidates = []

	for (const cookie of cookies) {
		if (!cookie?.value) {
			continue
		}

		const values = extractTokenCandidates(cookie.value)
		values.forEach((value) => {
			candidates.push({
				key: cookie.name,
				source: "cookie-api",
				value,
			})
		})
	}

	await ingestTokenCandidates(candidates, { pageUrl: "", pageTitle: "" })
}

async function ingestTokenCandidates(tokenCandidates, context = {}) {
	if (!Array.isArray(tokenCandidates) || !tokenCandidates.length) {
		return
	}

	for (const candidate of tokenCandidates) {
		if (!candidate?.value) {
			continue
		}

		await persistAuthToken(candidate.value, {
			source: candidate.source || "page-scan",
			key: candidate.key || "",
			url: context.pageUrl || "",
			title: context.pageTitle || "",
		})
	}
}

async function persistAuthToken(rawToken, meta = {}) {
	const token = normalizeToken(rawToken)
	if (!token) {
		return
	}

	const current = await chrome.storage.local.get([STORAGE_KEYS.AUTH_TOKEN, STORAGE_KEYS.AUTH_META])
	const currentToken = current[STORAGE_KEYS.AUTH_TOKEN] || ""
	const currentMeta = current[STORAGE_KEYS.AUTH_META] || null

	if (currentToken && scoreToken(currentToken) > scoreToken(token)) {
		return
	}

	if (currentToken === token && currentMeta?.source === meta.source) {
		return
	}

	await chrome.storage.local.set({
		[STORAGE_KEYS.AUTH_TOKEN]: token,
		[STORAGE_KEYS.AUTH_META]: {
			source: meta.source || "unknown",
			key: meta.key || "",
			url: meta.url || "",
			title: meta.title || "",
			updatedAt: new Date().toISOString(),
		},
	})
}

async function upsertQueryHistory(entry) {
	if (!entry?.query || typeof entry.query !== "object") {
		return
	}

	const storage = await chrome.storage.local.get(STORAGE_KEYS.QUERY_HISTORY)
	const queryHistory = Array.isArray(storage[STORAGE_KEYS.QUERY_HISTORY])
		? storage[STORAGE_KEYS.QUERY_HISTORY]
		: []

	const query = sanitizeCapturedQuery(entry.query)
	const signature = getQuerySignature(query)
	const index = queryHistory.findIndex((item) => item.signature === signature)

	const queryEntry = {
		id: `query_${signature}`,
		signature,
		name: sanitizeFileName(entry.name || "未命名查询"),
		query,
		source: entry.source || "unknown",
		pageUrl: entry.pageUrl || "",
		pageTitle: entry.pageTitle || "",
		capturedAt: new Date().toISOString(),
	}

	if (index >= 0) {
		queryHistory[index] = { ...queryHistory[index], ...queryEntry }
	} else {
		queryHistory.unshift(queryEntry)
	}

	const trimmedHistory = sortQueries(queryHistory).slice(0, QUERY_HISTORY_LIMIT)
	await chrome.storage.local.set({
		[STORAGE_KEYS.QUERY_HISTORY]: trimmedHistory,
	})
}

async function replaceQueryHistoryFromTemplates(templates, payload = {}, sender) {
	const storage = await chrome.storage.local.get(STORAGE_KEYS.QUERY_HISTORY)
	const existingQueryHistory = Array.isArray(storage[STORAGE_KEYS.QUERY_HISTORY])
		? storage[STORAGE_KEYS.QUERY_HISTORY]
		: []
	const nextQueryHistory = buildQueryHistoryFromTemplates(templates, payload, sender)

	if (isSameQueryHistory(existingQueryHistory, nextQueryHistory)) {
		return
	}

	await chrome.storage.local.set({
		[STORAGE_KEYS.QUERY_HISTORY]: nextQueryHistory,
	})
}

function buildQueryHistoryFromTemplates(templates, payload = {}, sender) {
	const nextQueryHistory = []

	for (const template of templates) {
		let templateQuery = template?.templateContent
		if (typeof templateQuery === "string") {
			try {
				templateQuery = JSON.parse(templateQuery)
			} catch (_error) {
				templateQuery = null
			}
		}

		if (!templateQuery || typeof templateQuery !== "object") {
			continue
		}

		const query = sanitizeCapturedQuery(templateQuery)
		const signature = getQuerySignature(query)

		nextQueryHistory.push({
			id: `query_${signature}`,
			signature,
			name: sanitizeFileName(template.templateName || guessQueryName(query, payload.pageTitle)),
			query,
			source: "template-response",
			pageUrl: payload.pageUrl || sender?.tab?.url || "",
			pageTitle: payload.pageTitle || sender?.tab?.title || "",
			capturedAt: template.updateTime || template.createTime || new Date().toISOString(),
		})
	}

	return sortQueries(nextQueryHistory).slice(0, QUERY_HISTORY_LIMIT)
}

function isSameQueryHistory(left, right) {
	const normalize = (list) =>
		(list || []).map((item) => ({
			signature: item.signature,
			name: item.name,
			source: item.source,
			capturedAt: item.capturedAt,
		}))
	const leftNormalized = normalize(left).sort((a, b) => a.signature.localeCompare(b.signature))
	const rightNormalized = normalize(right).sort((a, b) => a.signature.localeCompare(b.signature))

	return stableStringify(leftNormalized) === stableStringify(rightNormalized)
}

async function exportQueries(queryIds) {
	const storage = await chrome.storage.local.get([
		STORAGE_KEYS.AUTH_TOKEN,
		STORAGE_KEYS.QUERY_HISTORY,
		STORAGE_KEYS.EXPORT_STATE,
	])
	const authToken = storage[STORAGE_KEYS.AUTH_TOKEN] || ""
	const queryHistory = Array.isArray(storage[STORAGE_KEYS.QUERY_HISTORY])
		? storage[STORAGE_KEYS.QUERY_HISTORY]
		: []
	const exportState = await recoverExportStateIfNeeded(storage[STORAGE_KEYS.EXPORT_STATE])

	if (exportState.running) {
		throw new Error("已有导出任务正在执行")
	}

	if (!authToken) {
		throw new Error("没有抓到有效 token，请先登录页面并执行一次查询")
	}

	const selectedQueries = queryHistory.filter((item) => queryIds.includes(item.id))
	if (!selectedQueries.length) {
		throw new Error("请至少选择一个查询条件")
	}

	const exportSessionId = `export_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`
	activeExportSessionId = exportSessionId

	await setExportState({
		running: true,
		total: selectedQueries.length,
		completed: 0,
		currentName: "",
		message: "开始导出",
		error: "",
		updatedAt: new Date().toISOString(),
	})

	const results = []

	try {
		for (let index = 0; index < selectedQueries.length; index += 1) {
			const queryEntry = selectedQueries[index]
			await setExportState({
				running: true,
				total: selectedQueries.length,
				completed: index,
				currentName: queryEntry.name,
				message: `正在导出 ${queryEntry.name} (${index + 1}/${selectedQueries.length})`,
				error: "",
				progressPercent: 0,
				progressCurrent: 0,
				progressTarget: 0,
				updatedAt: new Date().toISOString(),
			})

			const result = await exportSingleQuery(queryEntry, authToken, async (progress) => {
				await setExportState({
					running: true,
					total: selectedQueries.length,
					completed: index,
					currentName: queryEntry.name,
					message: progress.message || `正在导出 ${queryEntry.name}`,
					error: "",
					progressPercent: progress.progressPercent,
					progressCurrent: progress.progressCurrent,
					progressTarget: progress.progressTarget,
					updatedAt: new Date().toISOString(),
				})
			})
			results.push(result)

			await setExportState({
				running: true,
				total: selectedQueries.length,
				completed: index + 1,
				currentName: queryEntry.name,
				message: `已完成 ${queryEntry.name}`,
				error: "",
				progressPercent: 100,
				progressCurrent: result.recordCount,
				progressTarget: result.recordCount,
				updatedAt: new Date().toISOString(),
			})
		}

		await setExportState({
			running: false,
			total: selectedQueries.length,
			completed: selectedQueries.length,
			currentName: "",
			message: `导出完成，共 ${selectedQueries.length} 个文件`,
			error: "",
			progressPercent: 100,
			progressCurrent: 0,
			progressTarget: 0,
			updatedAt: new Date().toISOString(),
		})

		return results
	} catch (error) {
		await setExportState({
			running: false,
			total: selectedQueries.length,
			completed: results.length,
			currentName: "",
			message: "导出失败",
			error: error.message,
			progressPercent: 0,
			progressCurrent: 0,
			progressTarget: 0,
			updatedAt: new Date().toISOString(),
		})
		throw error
	} finally {
		if (activeExportSessionId === exportSessionId) {
			activeExportSessionId = ""
		}
	}
}

async function exportSingleQuery(queryEntry, authToken, onProgress) {
	const exportQuery = buildExportQuery(queryEntry.query)
	const firstPage = await fetchPersonnelData(exportQuery, 1, authToken)
	const targetRecords = Math.min(Number(firstPage.total) || 0, MAX_RECORDS)

	let records = firstPage.records.slice(0, MAX_RECORDS)
	await reportExportProgress(onProgress, records.length, targetRecords, `正在获取人员信息: ${queryEntry.name}`)
	await enrichRecordsWithProjects(records, authToken)

	const totalPages = Math.max(1, Number(firstPage.totalPages) || 1)
	const maxPageToProcess = Math.min(totalPages, Math.ceil(MAX_RECORDS / exportQuery.pageSize))

	for (let pageNum = 2; pageNum <= maxPageToProcess && records.length < MAX_RECORDS; pageNum += 1) {
		const pageData = await fetchPersonnelData(exportQuery, pageNum, authToken)
		const remaining = MAX_RECORDS - records.length
		const pageRecords = pageData.records.slice(0, remaining)
		records = records.concat(pageRecords)
		await reportExportProgress(onProgress, records.length, targetRecords, `正在获取人员信息: ${queryEntry.name}`)
		await enrichRecordsWithProjects(pageRecords, authToken)
	}

	await reportExportProgress(onProgress, targetRecords || records.length, targetRecords || records.length, "正在生成 Excel")

	const workbook = buildWorkbook(records)
	const safeName = sanitizeFileName(queryEntry.name || guessQueryName(queryEntry.query, "查询"))
	const downloadName = `checkOwner/${formatDate(new Date())}/${safeName}-(${records.length}条).xlsx`
	const base64 = XLSX.write(workbook, { bookType: "xlsx", type: "base64" })

	await chrome.downloads.download({
		url: `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${base64}`,
		filename: downloadName,
		saveAs: false,
		conflictAction: "overwrite",
	})

	return {
		name: queryEntry.name,
		recordCount: records.length,
		fileName: downloadName,
	}
}

async function reportExportProgress(onProgress, current, target, message) {
	if (typeof onProgress !== "function") {
		return
	}

	const safeTarget = Number(target) || 0
	const safeCurrent = Math.min(Number(current) || 0, safeTarget || Number(current) || 0)
	const progressPercent = safeTarget ? Math.min(100, Math.round((safeCurrent / safeTarget) * 100)) : 100

	await onProgress({
		progressPercent,
		progressCurrent: safeCurrent,
		progressTarget: safeTarget,
		message,
	})
}

async function fetchPersonnelData(queryCondition, pageNum, authToken) {
	const queryToSend = {
		...queryCondition,
		pageNum,
	}

	const payload = await requestJson(PERSONNEL_ENDPOINT, queryToSend, authToken)
	const data = payload?.data || {}
	const records = Array.isArray(data.records) ? data.records : []
	const total = Number(data.total) || records.length
	const pageSize = Number(data.size) || Number(queryCondition.pageSize) || EXPORT_PAGE_SIZE
	const totalPages =
		Number(data.pages) ||
		Number(data.totalPages) ||
		(total && pageSize ? Math.ceil(total / pageSize) : 0) ||
		Number(data.pageNum) ||
		(records.length ? 1 : 0)

	return {
		records,
		total,
		totalPages,
		currentPage: Number(data.current) || pageNum,
		pageSize,
	}
}

async function requestJson(endpoint, body, authToken) {
	const response = await fetch(`${API_BASE_URL}${endpoint}`, {
		method: "POST",
		headers: {
			authorization: authToken,
			"content-type": "application/json;charset=UTF-8",
		},
		body: JSON.stringify(body),
	})

	const text = await response.text()
	let payload = null

	try {
		payload = text ? JSON.parse(text) : null
	} catch (_error) {
		throw new Error(`接口 ${endpoint} 返回了无法解析的内容`)
	}

	if (!response.ok) {
		throw new Error(payload?.message || `请求失败: ${response.status}`)
	}

	if (payload && payload.status === false) {
		throw new Error(payload.message || `接口 ${endpoint} 返回失败`)
	}

	return payload
}

async function enrichRecordsWithProjects(records, authToken) {
	for (let index = 0; index < records.length; index += 1) {
		const record = records[index]
		record.certificateInfo = extractCertificateInfo(record.dataTop)

		if (!record.uuid) {
			record.projectNames = []
			continue
		}

		try {
			record.projectNames = await fetchProjectNames(record.uuid, authToken)
		} catch (_error) {
			record.projectNames = []
		}
	}
}

async function fetchProjectNames(personId, authToken) {
	const payload = await requestJson(
		PROJECT_ENDPOINT,
		{
			personId,
			pageSize: 5,
			pageNum: 1,
		},
		authToken
	)

	const records = Array.isArray(payload?.data?.records) ? payload.data.records : []
	const uniqueNames = []

	for (const record of records) {
		if (!record?.projectName || uniqueNames.includes(record.projectName)) {
			continue
		}
		uniqueNames.push(record.projectName)
	}

	return uniqueNames
}

function extractCertificateInfo(dataTop) {
	if (!dataTop) {
		return null
	}

	return {
		certTime: dataTop.changeTime || "",
		certType: dataTop.typeName || "",
		companyInfo: dataTop.companyTwo || "",
		specialty: dataTop.registeredType || "",
	}
}

function buildWorkbook(records) {
	const workbook = XLSX.utils.book_new()
	const worksheet = {
		A1: { v: "姓名", t: "s" },
		B1: { v: "年龄", t: "s" },
		C1: { v: "身份证号", t: "s" },
		D1: { v: "业绩", t: "s" },
		E1: { v: "专业", t: "s" },
		F1: { v: "项目名称", t: "s" },
	}

	records.forEach((item, index) => {
		const rowNum = index + 2
		worksheet[`A${rowNum}`] = { v: item.name || "", t: "s" }
		worksheet[`B${rowNum}`] = { v: getAge(item), t: "n" }
		worksheet[`C${rowNum}`] = { v: item.idCard || "", t: "s" }
		worksheet[`D${rowNum}`] = { v: buildAchievementText(item), t: "s" }
		worksheet[`E${rowNum}`] = { v: buildProfessionalText(item), t: "s" }
		worksheet[`F${rowNum}`] = { v: buildProjectText(item), t: "s" }
	})

	const lastRow = Math.max(1, records.length + 1)
	worksheet["!ref"] = XLSX.utils.encode_range({
		s: { c: 0, r: 0 },
		e: { c: 5, r: lastRow - 1 },
	})
	worksheet["!cols"] = [
		{ wch: 12 },
		{ wch: 8 },
		{ wch: 18 },
		{ wch: 25 },
		{ wch: 50 },
		{ wch: 35 },
	]
	worksheet["!rows"] = Array.from({ length: lastRow }, (_, index) => ({
		hpt: index === 0 ? 30 : 100,
	}))

	XLSX.utils.book_append_sheet(workbook, worksheet, "人员数据")
	return workbook
}

function getAge(item) {
	let age = item.age
	if (!age && item.idCard) {
		try {
			const idCard = String(item.idCard).replace(/\*/g, "0")
			if (idCard.length >= 14) {
				const birthYear = parseInt(idCard.slice(6, 10), 10)
				if (birthYear > 1900 && birthYear < 2010) {
					age = new Date().getFullYear() - birthYear
				}
			}
		} catch (_error) {
			age = null
		}
	}

	if (!age || Number.isNaN(Number(age)) || age <= 0 || age > 100) {
		return 40
	}

	return Number(age)
}

function buildAchievementText(item) {
	const aAchieve = item.sikuLevelCountA || 0
	const bAchieve = item.sikuLevelCountB || 0
	const cAchieve = item.sikuLevelCountC || 0
	const dAchieve = item.sikuLevelCountD || 0
	return `A级业绩:${aAchieve}\nB级业绩:${bAchieve}\nC级业绩:${cAchieve}\nD级业绩:${dAchieve}`
}

function buildProfessionalText(item) {
	const certificateInfo = item.certificateInfo || {}
	return `时间:${certificateInfo.certTime || ""}\n状态:${certificateInfo.certType || ""}\n公司:${
		certificateInfo.companyInfo || ""
	}\n专业:${certificateInfo.specialty || ""}`
}

function buildProjectText(item) {
	const projectNames = Array.isArray(item.projectNames) ? [...item.projectNames] : []
	if (!projectNames.length && item.latestJson) {
		try {
			const latestData = typeof item.latestJson === "string" ? JSON.parse(item.latestJson) : item.latestJson
			if (latestData?.projectName) {
				projectNames.push(latestData.projectName)
			}
		} catch (_error) {
			// ignore parse failure
		}
	}

	return projectNames.join("\n")
}

function buildExportQuery(query) {
	return {
		...sanitizeCapturedQuery(query),
		pageNum: 1,
		pageSize: EXPORT_PAGE_SIZE,
		sortAsc: false,
		sort: 3,
	}
}

function sanitizeCapturedQuery(query) {
	return JSON.parse(JSON.stringify(query || {}))
}

function guessQueryName(query, pageTitle) {
	const certNames = Array.isArray(query?.arrayCertTypeNames) ? query.arrayCertTypeNames.filter(Boolean) : []
	if (certNames.length) {
		return certNames.join("_")
	}

	if (typeof pageTitle === "string" && pageTitle.trim()) {
		return pageTitle.trim().slice(0, 40)
	}

	return `查询_${formatDateTime(new Date())}`
}

function normalizeToken(rawToken) {
	const candidates = extractTokenCandidates(rawToken)
	if (!candidates.length) {
		return ""
	}

	return candidates.sort((left, right) => scoreToken(right) - scoreToken(left))[0]
}

function extractTokenCandidates(rawValue) {
	if (rawValue == null) {
		return []
	}

	const candidates = new Set()
	const stringValue = decodeSafely(typeof rawValue === "string" ? rawValue : JSON.stringify(rawValue))
	const prefixedMatches = stringValue.match(/(?:JwtA|Bearer)\s+[A-Za-z0-9._-]+/gi) || []
	prefixedMatches.forEach((item) => candidates.add(item.trim()))

	const directValue = stringValue.trim()
	if (/^[A-Za-z0-9-_]+\.[A-Za-z0-9-_]+\.[A-Za-z0-9-_]+$/.test(directValue)) {
		candidates.add(directValue)
	}

	try {
		const parsed = JSON.parse(directValue)
		collectTokensFromObject(parsed, candidates)
	} catch (_error) {
		// ignore
	}

	return Array.from(candidates)
}

function collectTokensFromObject(value, candidates, depth = 0) {
	if (depth > 4 || value == null) {
		return
	}

	if (typeof value === "string") {
		extractTokenCandidates(value).forEach((item) => candidates.add(item))
		return
	}

	if (Array.isArray(value)) {
		value.forEach((item) => collectTokensFromObject(item, candidates, depth + 1))
		return
	}

	if (typeof value === "object") {
		Object.entries(value).forEach(([key, item]) => {
			if (/token|auth|jwt/i.test(key)) {
				extractTokenCandidates(item).forEach((token) => candidates.add(token))
			}
			collectTokensFromObject(item, candidates, depth + 1)
		})
	}
}

function scoreToken(token) {
	let score = token.length
	if (/^JwtA\s+/i.test(token)) {
		score += 100
	}
	if (/^Bearer\s+/i.test(token)) {
		score += 80
	}
	if (token.split(".").length === 3) {
		score += 20
	}
	return score
}

function getQuerySignature(query) {
	return simpleHash(stableStringify(stripRuntimePaging(query)))
}

function stripRuntimePaging(query) {
	const cloned = sanitizeCapturedQuery(query)
	delete cloned.pageNum
	delete cloned.pageSize
	return cloned
}

function sortQueries(queryHistory) {
	return [...queryHistory].sort((left, right) => {
		const leftTime = new Date(left.capturedAt || 0).getTime()
		const rightTime = new Date(right.capturedAt || 0).getTime()
		return rightTime - leftTime
	})
}

function sanitizeFileName(fileName) {
	return String(fileName || "未命名查询")
		.replace(/[\\/:*?"<>|]/g, "_")
		.replace(/\s+/g, " ")
		.trim()
}

function stableStringify(value) {
	if (value == null || typeof value !== "object") {
		return JSON.stringify(value)
	}

	if (Array.isArray(value)) {
		return `[${value.map((item) => stableStringify(item)).join(",")}]`
	}

	const keys = Object.keys(value).sort()
	return `{${keys.map((key) => `${JSON.stringify(key)}:${stableStringify(value[key])}`).join(",")}}`
}

function simpleHash(text) {
	let hash = 0
	for (let index = 0; index < text.length; index += 1) {
		hash = (hash << 5) - hash + text.charCodeAt(index)
		hash |= 0
	}
	return Math.abs(hash).toString(36)
}

function buildIdleExportState() {
	return {
		running: false,
		total: 0,
		completed: 0,
		currentName: "",
		message: "等待导出",
		error: "",
		progressPercent: 0,
		progressCurrent: 0,
		progressTarget: 0,
		updatedAt: new Date().toISOString(),
	}
}

async function recoverExportStateIfNeeded(currentState) {
	let exportState = currentState
	if (typeof exportState === "undefined") {
		const storage = await chrome.storage.local.get(STORAGE_KEYS.EXPORT_STATE)
		exportState = storage[STORAGE_KEYS.EXPORT_STATE]
	}

	if (!exportState) {
		return buildIdleExportState()
	}

	if (!exportState.running || activeExportSessionId) {
		return exportState
	}

	const recoveredState = {
		...buildIdleExportState(),
		message: RECOVERED_EXPORT_MESSAGE,
		updatedAt: new Date().toISOString(),
	}

	await setExportState(recoveredState)
	return recoveredState
}

async function setExportState(nextState) {
	await chrome.storage.local.set({
		[STORAGE_KEYS.EXPORT_STATE]: nextState,
	})
}

function formatDate(date) {
	return [
		date.getFullYear(),
		String(date.getMonth() + 1).padStart(2, "0"),
		String(date.getDate()).padStart(2, "0"),
	].join("-")
}

function formatDateTime(date) {
	return [
		formatDate(date),
		`${String(date.getHours()).padStart(2, "0")}${String(date.getMinutes()).padStart(2, "0")}${String(
			date.getSeconds()
		).padStart(2, "0")}`,
	].join("_")
}

function decodeSafely(value) {
	try {
		return decodeURIComponent(value)
	} catch (_error) {
		return value
	}
}

function isSupportedPage(url) {
	if (!url) {
		return false
	}

	try {
		return /jiangongdata\.com$/i.test(new URL(url).hostname)
	} catch (_error) {
		return false
	}
}
