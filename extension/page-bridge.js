;(function bootstrapPageBridge() {
	if (window.__CHECKOWNER_PAGE_BRIDGE__) {
		return
	}

	window.__CHECKOWNER_PAGE_BRIDGE__ = true

	const SOURCE = "checkowner-page-bridge"
	const PERSONNEL_PATH = "/jian-butler-owner-biz/personnel/pagePersonnel"
	const TEMPLATE_LIST_PATH = "/jian-butler-owner-biz/searchTemplate/getOwnerSearchTemplateList"

	patchFetch()
	patchXMLHttpRequest()

	function patchFetch() {
		if (typeof window.fetch !== "function") {
			return
		}

		const nativeFetch = window.fetch
		window.fetch = function patchedFetch(input, init) {
			const requestUrl = resolveUrl(typeof input === "string" ? input : input?.url)
			const requestMethod = (init?.method || input?.method || "GET").toUpperCase()
			const requestHeaders = normalizeHeaders(init?.headers || input?.headers)

			maybeCaptureAuth(requestUrl, requestHeaders, "fetch")

			captureFetchBody(input, init)
				.then((bodyText) => {
					maybeCaptureQuery(requestUrl, requestMethod, requestHeaders, bodyText, "fetch")
				})
				.catch(() => {
					// body is optional
				})

			return nativeFetch.apply(this, arguments).then((response) => {
				maybeCaptureTemplateResponse(requestUrl, response.clone(), "fetch")
				return response
			})
		}
	}

	function patchXMLHttpRequest() {
		const originalOpen = XMLHttpRequest.prototype.open
		const originalSetRequestHeader = XMLHttpRequest.prototype.setRequestHeader
		const originalSend = XMLHttpRequest.prototype.send

		XMLHttpRequest.prototype.open = function patchedOpen(method, url) {
			this.__checkOwnerRequestMeta = {
				method: (method || "GET").toUpperCase(),
				url: resolveUrl(url),
				headers: {},
			}
			return originalOpen.apply(this, arguments)
		}

		XMLHttpRequest.prototype.setRequestHeader = function patchedSetRequestHeader(name, value) {
			if (this.__checkOwnerRequestMeta) {
				this.__checkOwnerRequestMeta.headers[String(name).toLowerCase()] = value
			}
			return originalSetRequestHeader.apply(this, arguments)
		}

		XMLHttpRequest.prototype.send = function patchedSend(body) {
			const meta = this.__checkOwnerRequestMeta || { method: "GET", url: location.href, headers: {} }

			maybeCaptureAuth(meta.url, meta.headers, "xhr")
			maybeCaptureQuery(meta.url, meta.method, meta.headers, body, "xhr")

			this.addEventListener("loadend", () => {
				maybeCaptureTemplateResponse(meta.url, this, "xhr")
			})

			return originalSend.apply(this, arguments)
		}
	}

	function maybeCaptureAuth(url, headers, transport) {
		if (!isJiangongUrl(url)) {
			return
		}

		const authorization = headers.authorization || headers.Authorization || headers.Authorization?.trim?.()
		if (!authorization) {
			return
		}

		emit("auth-captured", {
			authorization,
			source: transport,
			url,
		})
	}

	function maybeCaptureQuery(url, method, headers, body, transport) {
		if (method !== "POST" || !String(url).includes(PERSONNEL_PATH)) {
			return
		}

		const parsedBody = parseBody(body)
		if (!parsedBody || typeof parsedBody !== "object") {
			return
		}

		emit("query-captured", {
			query: parsedBody,
			source: transport,
			url,
			authorization: headers.authorization || headers.Authorization || "",
			hasAuthorization: Boolean(headers.authorization || headers.Authorization),
		})
	}

	function maybeCaptureTemplateResponse(url, responseLike, transport) {
		if (!String(url).includes(TEMPLATE_LIST_PATH)) {
			return
		}

		readResponseText(responseLike)
			.then((text) => {
				if (!text || text.length > 1024 * 1024) {
					return
				}

				const payload = parseJson(text)
				if (!payload) {
					return
				}

				emit("template-list-captured", {
					templates: extractTemplateItems(payload),
					source: transport,
					url,
				})
			})
			.catch(() => {
				// ignore response parsing issues
			})
	}

	function extractTemplateItems(payload) {
		if (Array.isArray(payload?.data)) {
			return payload.data
		}

		const queue = [payload]
		while (queue.length) {
			const current = queue.shift()
			if (!current) {
				continue
			}

			if (Array.isArray(current)) {
				if (current.every((item) => item && item.templateName && item.templateContent)) {
					return current
				}

				current.forEach((item) => queue.push(item))
				continue
			}

			if (typeof current === "object") {
				Object.values(current).forEach((value) => queue.push(value))
			}
		}

		return []
	}

	function parseJson(text) {
		try {
			return JSON.parse(text)
		} catch (_error) {
			return null
		}
	}

	function normalizeHeaders(headersLike) {
		const headers = {}

		if (!headersLike) {
			return headers
		}

		if (typeof Headers !== "undefined" && headersLike instanceof Headers) {
			headersLike.forEach((value, key) => {
				headers[key] = value
			})
			return headers
		}

		if (Array.isArray(headersLike)) {
			headersLike.forEach(([key, value]) => {
				headers[key] = value
			})
			return headers
		}

		if (typeof headersLike === "object") {
			Object.entries(headersLike).forEach(([key, value]) => {
				headers[key] = value
			})
		}

		return headers
	}

	async function captureFetchBody(input, init) {
		if (init?.body != null) {
			return stringifyBody(init.body)
		}

		if (typeof Request !== "undefined" && input instanceof Request) {
			try {
				return await input.clone().text()
			} catch (_error) {
				return ""
			}
		}

		return ""
	}

	function parseBody(body) {
		const bodyText = stringifyBody(body)
		if (!bodyText) {
			return null
		}

		try {
			return JSON.parse(bodyText)
		} catch (_error) {
			return null
		}
	}

	function stringifyBody(body) {
		if (body == null) {
			return ""
		}

		if (typeof body === "string") {
			return body
		}

		if (typeof URLSearchParams !== "undefined" && body instanceof URLSearchParams) {
			return body.toString()
		}

		if (typeof FormData !== "undefined" && body instanceof FormData) {
			const formObject = {}
			for (const [key, value] of body.entries()) {
				formObject[key] = value
			}
			return JSON.stringify(formObject)
		}

		if (typeof body === "object") {
			try {
				return JSON.stringify(body)
			} catch (_error) {
				return ""
			}
		}

		return String(body)
	}

	async function readResponseText(responseLike) {
		if (!responseLike) {
			return ""
		}

		if (typeof Response !== "undefined" && responseLike instanceof Response) {
			return responseLike.text()
		}

		if (typeof responseLike.responseText === "string") {
			return responseLike.responseText
		}

		return ""
	}

	function resolveUrl(inputUrl) {
		if (!inputUrl) {
			return location.href
		}

		try {
			return new URL(inputUrl, location.href).toString()
		} catch (_error) {
			return String(inputUrl)
		}
	}

	function isJiangongUrl(url) {
		try {
			return /jiangongdata\.com$/i.test(new URL(url).hostname)
		} catch (_error) {
			return false
		}
	}

	function emit(eventType, payload) {
		window.postMessage(
			{
				source: SOURCE,
				eventType,
				payload,
			},
			"*"
		)
	}
})()
