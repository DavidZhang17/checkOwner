#!/usr/bin/env node

const fs = require("fs")
const fsp = require("fs/promises")
const path = require("path")
const archiver = require("archiver")

const ROOT_DIR = path.resolve(__dirname, "..")
const DIST_DIR = path.join(ROOT_DIR, "dist")
const STAGING_DIR = path.join(DIST_DIR, ".staging")
const EXTENSION_DIR = path.join(ROOT_DIR, "extension")
const PACKAGE_JSON_PATH = path.join(ROOT_DIR, "package.json")
const MANIFEST_FILE = "manifest.json"
const SKIP_NAMES = new Set([".DS_Store"])

const TARGETS = {
	chrome: {
		extension: ".zip",
		updateManifest: (manifest) => manifest,
	},
	edge: {
		extension: ".zip",
		updateManifest: (manifest) => manifest,
	},
	firefox: {
		extension: ".xpi",
		updateManifest: (manifest) => ({
			...manifest,
			browser_specific_settings: {
				...(manifest.browser_specific_settings || {}),
				gecko: {
					id: "checkowner-export@local",
					strict_min_version: "121.0",
				},
			},
		}),
	},
}

main().catch((error) => {
	console.error(`[pack-extension] ${error.message}`)
	process.exitCode = 1
})

async function main() {
	const targetArg = process.argv[2] || "all"
	if (targetArg === "clean") {
		await cleanDist()
		console.log("[pack-extension] 已清理 dist 目录")
		return
	}

	const targets = resolveTargets(targetArg)
	const packageMeta = JSON.parse(await fsp.readFile(PACKAGE_JSON_PATH, "utf8"))
	const version = packageMeta.version || "0.0.0"

	await cleanDist()
	await fsp.mkdir(DIST_DIR, { recursive: true })

	for (const target of targets) {
		const targetConfig = TARGETS[target]
		const stagingPath = path.join(STAGING_DIR, target)
		const artifactName = `checkowner-${target}-v${version}${targetConfig.extension}`
		const artifactPath = path.join(DIST_DIR, artifactName)

		await copyDirectory(EXTENSION_DIR, stagingPath)
		await patchManifest(path.join(stagingPath, MANIFEST_FILE), targetConfig.updateManifest)
		await createArchive(stagingPath, artifactPath)

		console.log(`[pack-extension] 已生成 ${artifactName}`)
	}

	await fsp.rm(STAGING_DIR, { recursive: true, force: true })
}

function resolveTargets(targetArg) {
	if (targetArg === "all") {
		return Object.keys(TARGETS)
	}

	if (!TARGETS[targetArg]) {
		throw new Error(`不支持的目标平台: ${targetArg}`)
	}

	return [targetArg]
}

async function cleanDist() {
	await fsp.rm(DIST_DIR, { recursive: true, force: true })
}

async function copyDirectory(sourceDir, targetDir) {
	await fsp.rm(targetDir, { recursive: true, force: true })
	await fsp.mkdir(targetDir, { recursive: true })

	const entries = await fsp.readdir(sourceDir, { withFileTypes: true })
	for (const entry of entries) {
		if (SKIP_NAMES.has(entry.name)) {
			continue
		}

		const sourcePath = path.join(sourceDir, entry.name)
		const targetPath = path.join(targetDir, entry.name)

		if (entry.isDirectory()) {
			await copyDirectory(sourcePath, targetPath)
			continue
		}

		if (entry.isFile()) {
			await fsp.copyFile(sourcePath, targetPath)
		}
	}
}

async function patchManifest(manifestPath, updateManifest) {
	const manifest = JSON.parse(await fsp.readFile(manifestPath, "utf8"))
	const nextManifest = updateManifest(manifest)
	await fsp.writeFile(manifestPath, `${JSON.stringify(nextManifest, null, 2)}\n`, "utf8")
}

function createArchive(sourceDir, artifactPath) {
	return new Promise((resolve, reject) => {
		const output = fs.createWriteStream(artifactPath)
		const archive = archiver("zip", { zlib: { level: 9 } })

		output.on("close", resolve)
		output.on("error", reject)
		archive.on("error", reject)

		archive.pipe(output)
		archive.directory(sourceDir, false)
		archive.finalize()
	})
}
