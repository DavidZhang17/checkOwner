;(function bootstrapContentScript() {
	if (window.__CHECKOWNER_CONTENT_SCRIPT__) {
		rescanPageState()
		return
	}

	window.__CHECKOWNER_CONTENT_SCRIPT__ = true

	const BRIDGE_SOURCE = "checkowner-page-bridge"

	injectBridge()
	rescanPageState()

	window.addEventListener("message", handleBridgeMessage)

	chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
		if (message?.type === "rescanPage") {
			injectBridge()
			rescanPageState()
			sendResponse({ ok: true })
		}
	})

	function handleBridgeMessage(event) {
		if (event.source !== window) {
			return
		}

		const data = event.data
		if (!data || data.source !== BRIDGE_SOURCE || !data.eventType) {
			return
		}

		chrome.runtime.sendMessage({
			type: "bridge-event",
			eventType: data.eventType,
			payload: {
				...data.payload,
				pageUrl: location.href,
				pageTitle: document.title || "",
			},
		})
	}

	function injectBridge(retryCount = 0) {
		if (document.getElementById("checkowner-page-bridge")) {
			return
		}

		const root = document.documentElement || document.head || document.body
		if (!root) {
			if (retryCount < 20) {
				setTimeout(() => injectBridge(retryCount + 1), 50)
			}
			return
		}

		const script = document.createElement("script")
		script.id = "checkowner-page-bridge"
		script.src = chrome.runtime.getURL("page-bridge.js")
		script.async = false
		script.onload = () => script.remove()
		script.onerror = () => script.remove()
		root.appendChild(script)
	}

	function rescanPageState() {
		const tokenCandidates = []
		pushStorageCandidates(window.localStorage, "localStorage", tokenCandidates)
		pushStorageCandidates(window.sessionStorage, "sessionStorage", tokenCandidates)
		pushStorageCandidatesFromCookies(tokenCandidates)

		chrome.runtime.sendMessage({
			type: "bridge-event",
			eventType: "page-scan",
			payload: {
				pageUrl: location.href,
				pageTitle: document.title || "",
				tokenCandidates,
			},
		})
	}

	function pushStorageCandidates(storageLike, storageName, tokenCandidates) {
		if (!storageLike) {
			return
		}

		try {
			for (let index = 0; index < storageLike.length; index += 1) {
				const key = storageLike.key(index)
				if (!key) {
					continue
				}

				const rawValue = storageLike.getItem(key)
				if (!rawValue) {
					continue
				}

				extractCandidateValues(rawValue).forEach((value) => {
					tokenCandidates.push({
						key,
						source: storageName,
						value,
					})
				})
			}
		} catch (error) {
			console.warn("[CheckOwner] 读取页面存储失败", error)
		}
	}

	function pushStorageCandidatesFromCookies(tokenCandidates) {
		try {
			const cookieEntries = document.cookie
				.split(";")
				.map((item) => item.trim())
				.filter(Boolean)

			cookieEntries.forEach((entry) => {
				const [rawKey, ...rawValueParts] = entry.split("=")
				const rawValue = rawValueParts.join("=")
				extractCandidateValues(rawValue).forEach((value) => {
					tokenCandidates.push({
						key: rawKey,
						source: "cookie",
						value,
					})
				})
			})
		} catch (error) {
			console.warn("[CheckOwner] 读取 cookie 失败", error)
		}
	}

	function extractCandidateValues(rawValue) {
		const candidates = new Set()
		const stringValue = decodeSafely(rawValue)
		const compact = typeof stringValue === "string" ? stringValue.trim() : ""

		if (!compact) {
			return []
		}

		collectTokenLikes(compact, candidates)

		try {
			const parsed = JSON.parse(compact)
			scanObjectForTokens(parsed, candidates)
		} catch (_error) {
			// ignore non-json payloads
		}

		return Array.from(candidates)
	}

	function scanObjectForTokens(value, candidates, depth = 0) {
		if (depth > 4 || value == null) {
			return
		}

		if (typeof value === "string") {
			collectTokenLikes(value, candidates)
			return
		}

		if (Array.isArray(value)) {
			value.forEach((item) => scanObjectForTokens(item, candidates, depth + 1))
			return
		}

		if (typeof value === "object") {
			Object.entries(value).forEach(([key, item]) => {
				if (typeof item === "string" && /token|auth|jwt/i.test(key)) {
					collectTokenLikes(item, candidates)
				}
				scanObjectForTokens(item, candidates, depth + 1)
			})
		}
	}

	function collectTokenLikes(text, candidates) {
		if (!text) {
			return
		}

		const decoded = decodeSafely(text)
		const prefixedMatches = decoded.match(/(?:JwtA|Bearer)\s+[A-Za-z0-9._-]+/gi) || []
		prefixedMatches.forEach((item) => candidates.add(item.trim()))

		const jwtMatches = decoded.match(/[A-Za-z0-9-_]+\.[A-Za-z0-9-_]+\.[A-Za-z0-9-_]+/g) || []
		jwtMatches.forEach((item) => {
			if (item.split(".").length === 3) {
				candidates.add(item.trim())
			}
		})
	}

	function decodeSafely(value) {
		if (typeof value !== "string") {
			return ""
		}

		try {
			return decodeURIComponent(value)
		} catch (_error) {
			return value
		}
	}
})()
