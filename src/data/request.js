const axios = require("axios")
module.exports = axios.create({
	baseURL: "https://gwowner.jiangongdata.com",
	headers: {
		authorization:
			"JwtA eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJpc3N1ZXIiLCJidXNpbmVzc0lkIjoiT1dORVIxMDg5NzY1MTM5MjI0MjE1NTUyIiwiaWF0IjoxNzY3NDk0NzM4fQ.uujLKuA1hNavk2UJUaCmHMFes5nXwY_WjUtqRV4pL8o",
	},
})
