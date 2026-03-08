# 查业主-Excel导出

这个仓库现在只保留浏览器扩展方案，不再包含旧的 Node.js 抓取和导出流程。

## 目录

- `extension/`: 扩展源码
- `scripts/package-extension.js`: 浏览器扩展打包脚本

## 本地加载

1. 打开 Chrome，进入 `chrome://extensions`
2. 开启右上角“开发者模式”
3. 点击“加载已解压的扩展程序”
4. 选择当前项目下的 `extension/` 目录

扩展会在 `https://owner.jiangongdata.com/registerPerson` 页面自动显示右侧浮层面板。

## 打包

先安装开发依赖：

```bash
npm install
```

然后按目标浏览器打包：

```bash
npm run pack:chrome
npm run pack:edge
npm run pack:firefox
```

一次性打全部包：

```bash
npm run pack
```

生成结果会输出到 `dist/`：

- `checkowner-chrome-v<version>.zip`
- `checkowner-edge-v<version>.zip`
- `checkowner-firefox-v<version>.xpi`

## 使用流程

1. 登录建工数据页面
2. 打开 `https://owner.jiangongdata.com/registerPerson`
3. 页面右侧会自动出现“查业主-Excel导出”面板
4. 等页面调用模板列表接口后，条件会自动出现在扩展列表里
5. 勾选要导出的条件，点击“导出选中”

下载文件会进入浏览器默认下载目录下的 `checkOwner/YYYY-MM-DD/`。
