# CheckOwner 浏览器扩展

这个目录是浏览器扩展源码，可直接加载到 Chrome / Edge，也可以打包成 Firefox 扩展包。

## 功能

- 自动从 `*.jiangongdata.com` 页面抓取 `authorization` / JWT
- 条件列表来自页面实际调用的模板列表接口返回
- 在 `https://owner.jiangongdata.com/registerPerson` 页面自动注入右侧工具面板
- 直接在扩展里导出 Excel

## 加载到 Chrome / Edge

1. 打开 Chrome，进入 `chrome://extensions`
2. 开启右上角“开发者模式”
3. 点击“加载已解压的扩展程序”
4. 选择当前项目下的 `extension/` 目录

Edge 的加载方式相同，对应页面是 `edge://extensions`。

## 打包

在项目根目录执行：

```bash
npm install
npm run pack:chrome
npm run pack:edge
npm run pack:firefox
```

打包文件输出到根目录下的 `dist/`。

## 使用

1. 打开建工数据相关页面并完成登录
2. 打开 `https://owner.jiangongdata.com/registerPerson`
3. 页面右侧会自动出现“查业主-Excel导出”工具面板
4. 等页面调用模板列表接口后，条件会自动进入列表
5. 如果还没抓到数据，点一次“重新扫描”
6. 勾选要导出的条件，点击“导出选中”

下载的文件会进入浏览器默认下载目录下的 `checkOwner/YYYY-MM-DD/`。

## 说明

- 扩展默认每个查询最多导出 `5000` 条
- `token` 主要通过页面请求头、页面存储和 cookie 自动抓取
- Chrome 不允许扩展在标签页打开时自动弹出工具栏 popup，所以这里用的是“页面内自动浮层”方案
- 如果你是在安装扩展之前就已经打开了页面，第一次使用时最好刷新一次页面，或在面板里点“重新扫描”
