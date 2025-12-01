const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");
const axios = require("axios"); // 添加HTTP请求库
const request = axios.create({
  baseURL: "https://gwowner.jiangongdata.com",
  headers: {
    authorization:
      "JwtA eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJpc3N1ZXIiLCJidXNpbmVzc0lkIjoiT1dORVIxMDg5NzY1MTM5MjI0MjE1NTUyIiwiaWF0IjoxNzYyNzQxNjYyfQ.ylikBtB7IlOSmDBkF8_6CguCF6IszEG0vrGub8Gha18",
  },
});

// 读取查询条件配置文件
const queriesFilePath = path.join(__dirname, "queries.json");
let fetchPersonnelDataQueries = [];

try {
  const queriesContent = fs.readFileSync(queriesFilePath, "utf-8");
  fetchPersonnelDataQueries = JSON.parse(queriesContent);
  console.log(
    `✓ 成功加载查询配置文件，共 ${fetchPersonnelDataQueries.length} 个查询条件`
  );
} catch (error) {
  console.error(`❌ 无法读取查询配置文件 ${queriesFilePath}:`, error.message);
  console.error("请确保 queries.json 文件存在且格式正确");
  process.exit(1);
}

// 确保 files 文件夹存在
const filesDir = path.join(__dirname, "files");
if (!fs.existsSync(filesDir)) {
  fs.mkdirSync(filesDir, { recursive: true });
  console.log(`✓ 创建 files 文件夹: ${filesDir}`);
}

// 创建以当前日期为名称的子文件夹
const today = new Date();
const dateFolder = `${today.getFullYear()}-${String(
  today.getMonth() + 1
).padStart(2, "0")}-${String(today.getDate()).padStart(2, "0")}`;
const outputDir = path.join(filesDir, dateFolder);
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
  console.log(`✓ 创建日期文件夹: ${dateFolder}`);
} else {
  console.log(`✓ 使用已存在的日期文件夹: ${dateFolder}`);
}

// ==================== 配置说明 ====================
// 查询条件配置在 queries.json 文件中
// 每个条件包含：
// - fileName: 生成的Excel文件名（不含.xlsx后缀）
// - query: 查询条件对象
//   - pageNum: 起始页码（首次运行设置为1，断点续传时设置为中断的下一页）
//   - pageSize: 每页大小（建议值：1000-5000）
//
// 生成的Excel文件将保存在 files/YYYY-MM-DD/ 文件夹中（按日期自动分类）

// 使用async/await执行异步操作
(async () => {
  try {
    console.log(`\n==================== 开始处理 ====================`);
    console.log(`共有 ${fetchPersonnelDataQueries.length} 个查询条件需要处理`);
    console.log(`输出目录: ${outputDir}`);
    console.log("================================================\n");

    // 循环处理每个查询条件
    for (let i = 0; i < fetchPersonnelDataQueries.length; i++) {
      const queryConfig = fetchPersonnelDataQueries[i];
      const { fileName, query } = queryConfig;

      console.log(
        `\n[${i + 1}/${fetchPersonnelDataQueries.length}] 开始处理: ${fileName}`
      );
      console.log(`起始页码: ${query.pageNum}`);
      console.log(`每页大小: ${query.pageSize}`);
      console.log(
        `数据将写入Excel第 ${(query.pageNum - 1) * query.pageSize + 2} 行开始`
      );
      if (query.pageNum > 1) {
        console.log(`⚠️  检测到断点续传模式，请确保输出文件路径正确！`);
        console.log(`已跳过前 ${(query.pageNum - 1) * query.pageSize} 条记录`);
      }
      console.log("-------------------------------------------");

      // 执行转换
      await convertJsonToExcel(outputDir, query, fileName);

      console.log(
        `✓ [${i + 1}/${fetchPersonnelDataQueries.length}] 完成: ${fileName}`
      );

      // 如果不是最后一个，等待一段时间
      if (i < fetchPersonnelDataQueries.length - 1) {
        console.log(`\n休息 5 秒后继续下一个查询...\n`);
        await new Promise((resolve) => setTimeout(resolve, 5000));
      }
    }

    console.log(`\n==================== 全部完成 ====================`);
    console.log(`已成功处理 ${fetchPersonnelDataQueries.length} 个查询条件`);
    console.log(`所有文件已保存到: ${outputDir}`);
    console.log("================================================\n");
  } catch (error) {
    console.error("转换过程中出现错误:", error);
    process.exit(1);
  }
})();

/**
 * 通过API获取指定人员的项目名称列表
 * @param {string} personId 人员ID
 * @returns {Promise<Array<string>>} 项目名称列表
 */
async function fetchProjectNames(personId) {
  try {
    const response = await request.post(
      "/jian-butler-owner-biz/person/achievement/pageProjectWinningByPerson",
      {
        personId: personId,
        pageSize: 5,
        pageNum: 1,
      }
    );

    if (
      response.data &&
      response.data.status &&
      response.data.data &&
      response.data.data.records
    ) {
      // 提取projectName字段，并去重
      const projectNames = response.data.data.records
        .map((record) => record.projectName)
        .filter((name, index, self) => name && self.indexOf(name) === index);
      return projectNames;
    }
    return [];
  } catch (error) {
    console.warn(`获取项目名称失败 (${personId}): ${error.message}`);
    return [];
  }
}

/**
 * 从人员数据的dataTop字段提取证书信息
 * @param {Object} dataTop 人员数据中的dataTop字段
 * @returns {Object|null} 证书信息对象
 */
function extractCertificateInfo(dataTop) {
  if (!dataTop) {
    return null;
  }

  // 时间
  const certTime = dataTop.changeTime || "";

  // 类型状态
  const certType = dataTop.typeName || "";

  // 公司信息
  let companyInfo = "";
  if (dataTop.companyTwo) {
    companyInfo = dataTop.companyTwo;
  }

  // 专业
  const specialty = dataTop.registeredType || "";

  return {
    certTime,
    certType,
    companyInfo,
    specialty,
  };
}

/**
 * 从API获取人员数据
 * @param {Object} queryCondition 查询条件对象
 * @param {number} pageNum 页码
 * @returns {Promise<Object>} 人员数据列表
 */
async function fetchPersonnelData(queryCondition, pageNum) {
  try {
    console.log(`正在获取第 ${pageNum} 页数据...`);
    // 更新查询参数
    const queryToSend = {
      ...queryCondition,
      pageNum: pageNum,
    };

    const response = await request.post(
      "/jian-butler-owner-biz/personnel/pagePersonnel",
      queryToSend
    );

    if (
      response.data &&
      response.data.status &&
      response.data.data &&
      response.data.data.records
    ) {
      return {
        records: response.data.data.records,
        total: response.data.data.total || 0,
        totalPages: response.data.data.pageNum || 1,
        currentPage: response.data.data.current || pageNum,
        pageSize: response.data.data.size || queryCondition.pageSize,
      };
    }
    return {
      records: [],
      total: 0,
      totalPages: 0,
      currentPage: 0,
      pageSize: 0,
    };
  } catch (error) {
    console.error(`获取人员数据失败: ${error.message}`);
    throw error;
  }
}

/**
 * 为人员数据批量获取项目名称和提取证书信息
 * @param {Array} records 人员记录数组
 * @returns {Promise<void>}
 */
async function enrichRecordsWithProjects(records) {
  console.log(`开始为 ${records.length} 条记录获取项目名称和提取证书信息...`);
  for (let i = 0; i < records.length; i++) {
    const record = records[i];

    // 从dataTop字段提取证书信息
    record.certificateInfo = extractCertificateInfo(record.dataTop);

    if (record.uuid) {
      try {
        // 获取项目名称
        const projectNames = await fetchProjectNames(record.uuid);
        record.projectNames = projectNames;

        if ((i + 1) % 10 === 0) {
          console.log(`  已处理 ${i + 1}/${records.length} 条记录`);
        }
      } catch (e) {
        console.warn(`获取 ${record.name} 的项目信息失败: ${e.message}`);
        record.projectNames = [];
      }
    } else {
      record.projectNames = [];
    }
  }
  console.log(`项目名称获取和证书信息提取完成`);
}

/**
 * 将数据追加到Excel文件
 * @param {Array} records 要追加的记录
 * @param {string} outputPath Excel文件路径
 * @param {number} pageNum 当前页码
 * @param {number} pageSize 每页大小
 */
async function appendToExcel(records, outputPath, pageNum, pageSize) {
  let workbook;
  let worksheet;

  // 根据 pageNum 和 pageSize 计算起始行
  // 公式: (pageNum - 1) * pageSize + 2 （第1行是标题，第2行开始是数据）
  const startRow = (pageNum - 1) * pageSize + 2;

  const fileExists = fs.existsSync(outputPath);

  if (!fileExists || pageNum === 1) {
    // 创建新工作簿和工作表
    console.log(`创建新的Excel文件: ${outputPath}`);
    workbook = XLSX.utils.book_new();
    worksheet = {
      A1: { v: "姓名", t: "s" },
      B1: { v: "年龄", t: "s" },
      C1: { v: "身份证号", t: "s" },
      D1: { v: "业绩", t: "s" },
      E1: { v: "专业", t: "s" },
      F1: { v: "项目名称", t: "s" },
    };
  } else {
    // 读取现有文件
    console.log(`追加数据到现有文件（从第 ${startRow} 行开始）...`);
    workbook = XLSX.readFile(outputPath);
    worksheet = workbook.Sheets["人员数据"];
  }

  // 添加数据行
  console.log(`开始写入 ${records.length} 条记录到Excel...`);
  for (let index = 0; index < records.length; index++) {
    const item = records[index];
    const rowNum = startRow + index;

    // 1. 姓名
    worksheet[`A${rowNum}`] = { v: item.name || "", t: "s" };

    // 2. 年龄
    let age = item.age;
    if (!age && item.idCard) {
      try {
        const idCard = item.idCard.toString().replace(/\*/g, "0");
        if (idCard.length >= 14) {
          const birthYear = parseInt(idCard.substring(6, 10));
          if (birthYear && birthYear > 1900 && birthYear < 2010) {
            const currentYear = new Date().getFullYear();
            age = currentYear - birthYear;
          }
        }
      } catch (e) {
        /* 计算失败 */
      }
    }
    if (!age || isNaN(age) || age <= 0 || age > 100) {
      age = 40;
    }
    worksheet[`B${rowNum}`] = { v: age, t: "n" };

    // 3. 身份证号
    worksheet[`C${rowNum}`] = { v: item.idCard || "", t: "s" };

    // 4. 业绩
    const aAchieve = item.sikuLevelCountA || 0;
    const bAchieve = item.sikuLevelCountB || 0;
    const cAchieve = item.sikuLevelCountC || 0;
    const dAchieve = item.sikuLevelCountD || 0;
    const achieveText = `A级业绩:${aAchieve}\nB级业绩:${bAchieve}\nC级业绩:${cAchieve}\nD级业绩:${dAchieve}`;
    worksheet[`D${rowNum}`] = { v: achieveText, t: "s" };

    // 5. 专业信息 - 从dataTop字段提取的证书信息
    let certTime = "";
    let certType = "";
    let companyInfo = "";
    let specialty = "";

    if (item.certificateInfo) {
      certTime = item.certificateInfo.certTime || "";
      certType = item.certificateInfo.certType || "";
      companyInfo = item.certificateInfo.companyInfo || "";
      specialty = item.certificateInfo.specialty || "";
    }

    const professionalText = `时间:${certTime}\n状态:${certType}\n公司:${companyInfo}\n专业:${specialty}`;
    worksheet[`E${rowNum}`] = { v: professionalText, t: "s" };

    // 6. 项目名称
    let projectNames = item.projectNames || [];
    if (projectNames.length === 0 && item.latestJson) {
      try {
        let latestData = item.latestJson;
        if (typeof latestData === "string") {
          latestData = JSON.parse(latestData);
        }
        if (latestData.projectName) {
          projectNames.push(latestData.projectName);
        }
      } catch (e) {
        /* 解析失败 */
      }
    }
    const projectNamesText = projectNames.join("\n");
    worksheet[`F${rowNum}`] = { v: projectNamesText, t: "s" };
  }

  // 更新工作表范围
  const lastRow = startRow + records.length - 1;

  // 如果是追加，需要考虑原有范围
  let finalLastRow = lastRow;
  if (fileExists && pageNum > 1) {
    const existingRange = XLSX.utils.decode_range(worksheet["!ref"]);
    finalLastRow = Math.max(existingRange.e.r, lastRow);
  }

  const range = { s: { c: 0, r: 0 }, e: { c: 5, r: finalLastRow } };
  worksheet["!ref"] = XLSX.utils.encode_range(range);

  // 设置列宽
  worksheet["!cols"] = [
    { wch: 12 }, // 姓名
    { wch: 8 }, // 年龄
    { wch: 18 }, // 身份证号
    { wch: 25 }, // 业绩
    { wch: 50 }, // 专业
    { wch: 35 }, // 项目名称
  ];

  // 设置行高（只设置新增的行）
  if (!worksheet["!rows"]) {
    worksheet["!rows"] = [];
  }
  for (let i = 0; i <= finalLastRow; i++) {
    if (!worksheet["!rows"][i]) {
      worksheet["!rows"][i] = { hpt: i === 0 ? 30 : 100 };
    }
  }

  // 保存文件
  if (!fileExists || pageNum === 1) {
    XLSX.utils.book_append_sheet(workbook, worksheet, "人员数据");
  }
  XLSX.writeFile(workbook, outputPath);
  console.log(
    `✓ 已写入 ${records.length} 条记录（第 ${startRow} - ${lastRow} 行），当前文件共 ${finalLastRow} 行数据\n`
  );
}

/**
 * 获取所有人员数据并实时写入Excel
 * @param {string} outputPath Excel文件路径
 * @param {Object} queryCondition 查询条件对象
 */
async function fetchAndSaveData(outputPath, queryCondition, fileName) {
  const MAX_RECORDS = 5000; // 最大记录数限制
  let totalRecordsProcessed = 0;

  const startPageNum = queryCondition.pageNum;
  const pageSize = queryCondition.pageSize;

  try {
    // 获取第一页数据以获取总页数和总记录数
    console.log(`\n========== 开始获取第 ${startPageNum} 页数据 ==========`);
    const firstPage = await fetchPersonnelData(queryCondition, startPageNum);
    console.log(
      `第 ${startPageNum} 页: 获取到 ${firstPage.records.length} 条人员记录`
    );

    // 为第一页数据获取项目名称
    await enrichRecordsWithProjects(firstPage.records);

    // 生成输出文件路径（保存到日期文件夹）
    const outputPath2 = path.join(
      outputPath,
      `${fileName}-(${firstPage.records.length}条).xlsx`
    );

    // 写入Excel
    await appendToExcel(firstPage.records, outputPath2, startPageNum, pageSize);
    totalRecordsProcessed += firstPage.records.length;

    const totalPages = firstPage.totalPages;
    const totalRecords = firstPage.total;
    console.log(`总共 ${totalPages} 页，共 ${totalRecords} 条记录`);

    // 计算已完成的记录数（基于起始页码）
    const alreadyProcessed = (startPageNum - 1) * pageSize;
    console.log(
      `从第 ${startPageNum} 页继续，之前已完成 ${alreadyProcessed} 条记录`
    );

    // 检查是否超过最大记录数限制
    if (totalRecords > MAX_RECORDS) {
      console.log(
        `\n⚠️  检测到总记录数 ${totalRecords} 超过限制 ${MAX_RECORDS}`
      );
      console.log(`将只处理前 ${MAX_RECORDS} 条记录`);
    }

    // 计算需要处理的最大页数
    const maxPageToProcess = Math.min(
      totalPages,
      Math.ceil(MAX_RECORDS / pageSize)
    );

    console.log(`休息 1 分钟后继续...\n`);
    // await new Promise((resolve) => setTimeout(resolve, 60000));

    // 获取剩余页数据
    for (
      let pageNum = startPageNum + 1;
      pageNum <= maxPageToProcess;
      pageNum++
    ) {
      // 检查是否已达到5000条限制
      const currentTotal = alreadyProcessed + totalRecordsProcessed;
      if (currentTotal >= MAX_RECORDS) {
        console.log(`\n✓ 已达到最大记录数限制 ${MAX_RECORDS}，停止获取`);
        break;
      }

      console.log(`\n========== 开始获取第 ${pageNum} 页数据 ==========`);
      const pageData = await fetchPersonnelData(queryCondition, pageNum);
      console.log(
        `第 ${pageNum} 页: 获取到 ${pageData.records.length} 条人员记录`
      );

      // 如果当前页会超过5000限制，只取需要的记录数
      let recordsToProcess = pageData.records;
      if (currentTotal + pageData.records.length > MAX_RECORDS) {
        const remainingRecords = MAX_RECORDS - currentTotal;
        recordsToProcess = pageData.records.slice(0, remainingRecords);
        console.log(`⚠️  截取前 ${remainingRecords} 条记录以满足限制`);
      }

      // 为当前页数据获取项目名称
      await enrichRecordsWithProjects(recordsToProcess);

      // 追加写入Excel
      await appendToExcel(recordsToProcess, outputPath2, pageNum, pageSize);
      totalRecordsProcessed += recordsToProcess.length;

      const newTotal = alreadyProcessed + totalRecordsProcessed;
      console.log(
        `进度: ${newTotal}/${Math.min(totalRecords, MAX_RECORDS)} (${(
          (newTotal / Math.min(totalRecords, MAX_RECORDS)) *
          100
        ).toFixed(2)}%)`
      );

      // 如果已经达到限制，跳出循环
      if (newTotal >= MAX_RECORDS) {
        console.log(`\n✓ 已达到最大记录数限制 ${MAX_RECORDS}，停止获取`);
        break;
      }

      // console.log(`休息 1 分钟后继续...\n`);
      // await new Promise((resolve) => setTimeout(resolve, 60000));
    }

    console.log(`\n========== 数据获取完成 ==========`);
    console.log(`本次处理 ${totalRecordsProcessed} 条记录`);
    console.log(`总计 ${alreadyProcessed + totalRecordsProcessed} 条记录`);
    console.log(`文件保存在: ${outputPath2}`);
  } catch (error) {
    console.error(`处理数据时出错: ${error.message}`);
    console.error(`已成功保存 ${totalRecordsProcessed} 条记录到文件`);
    console.error(`如需继续，请修改 query.pageNum 继续运行`);
    throw error;
  }
}

/**
 * 将JSON数据转换为Excel文件
 * 使用增量写入方式，每获取一页数据就写入Excel
 * @param {string} outputPath 输出文件路径
 * @param {Object} queryCondition 查询条件对象
 */
async function convertJsonToExcel(outputPath, queryCondition, fileName) {
  try {
    console.log("开始从API获取人员数据并实时写入Excel...");
    console.log(`输出文件: ${outputPath}`);
    console.log(`起始页码: ${queryCondition.pageNum}`);
    console.log(`每页大小: ${queryCondition.pageSize}`);
    console.log(
      `数据将从第 ${
        (queryCondition.pageNum - 1) * queryCondition.pageSize + 2
      } 行开始写入`
    );

    // 获取数据并实时写入Excel
    await fetchAndSaveData(outputPath, queryCondition, fileName);

    console.log(`\n转换完成！文件已保存到: ${outputPath}`);
    return outputPath;
  } catch (error) {
    console.error(`转换失败: ${error.message}`);
    throw error;
  }
}

/**
 * 创建工作表
 */
async function createWorksheet(data) {
  const ws = {};

  // 创建标题行
  ws["A1"] = { v: "姓名", t: "s" };
  ws["B1"] = { v: "年龄", t: "s" };
  ws["C1"] = { v: "身份证号", t: "s" };
  ws["D1"] = { v: "业绩", t: "s" };
  ws["E1"] = { v: "专业", t: "s" };
  ws["F1"] = { v: "项目名称", t: "s" };

  // 添加数据行
  console.log(`开始生成 Excel 工作表，共 ${data.length} 条记录...`);
  for (let index = 0; index < data.length; index++) {
    const item = data[index];
    const rowNum = index + 2; // 从第2行开始

    // 进度提示
    if ((index + 1) % 1000 === 0) {
      console.log(`已处理 ${index + 1}/${data.length} 条记录...`);
    }

    // 1. 姓名
    ws[`A${rowNum}`] = { v: item.name || "", t: "s" };

    // 2. 年龄 - 直接使用age字段
    let age = item.age;

    // 如果age字段不存在或无效，则尝试从身份证计算
    if (!age && item.idCard) {
      try {
        const idCard = item.idCard.toString().replace(/\*/g, "0");
        if (idCard.length >= 14) {
          const birthYear = parseInt(idCard.substring(6, 10));
          if (birthYear && birthYear > 1900 && birthYear < 2010) {
            const currentYear = new Date().getFullYear();
            age = currentYear - birthYear;
          }
        }
      } catch (e) {
        /* 计算失败 */
      }
    }

    // 如果仍然没有有效年龄，使用默认值
    if (!age || isNaN(age) || age <= 0 || age > 100) {
      age = 40; // 默认年龄
    }
    ws[`B${rowNum}`] = { v: age, t: "n" };

    // 3. 身份证号
    ws[`C${rowNum}`] = { v: item.idCard || "", t: "s" };

    // 4. 业绩
    const aAchieve = item.sikuLevelCountA || 0;
    const bAchieve = item.sikuLevelCountB || 0;
    const cAchieve = item.sikuLevelCountC || 0;
    const dAchieve = item.sikuLevelCountD || 0;
    const achieveText = `A级业绩:${aAchieve}\nB级业绩:${bAchieve}\nC级业绩:${cAchieve}\nD级业绩:${dAchieve}`;
    ws[`D${rowNum}`] = { v: achieveText, t: "s" };

    // 5. 专业信息 - 从dataTop字段提取的证书信息
    let certTime = "";
    let certType = "";
    let companyInfo = "";
    let specialty = "";

    if (item.certificateInfo) {
      certTime = item.certificateInfo.certTime || "";
      certType = item.certificateInfo.certType || "";
      companyInfo = item.certificateInfo.companyInfo || "";
      specialty = item.certificateInfo.specialty || "";
    }

    const professionalText = `时间:${certTime}\n状态:${certType}\n公司:${companyInfo}\n专业:${specialty}`;
    ws[`E${rowNum}`] = { v: professionalText, t: "s" };

    // 6. 项目名称 - 使用已获取的项目名称
    let projectNames = item.projectNames || [];

    // 如果没有预先获取的项目名称，尝试使用latestJson
    if (projectNames.length === 0 && item.latestJson) {
      try {
        let latestData = item.latestJson;
        if (typeof latestData === "string") {
          latestData = JSON.parse(latestData);
        }
        if (latestData.projectName) {
          projectNames.push(latestData.projectName);
        }
      } catch (e) {
        console.warn(`解析latestJson失败: ${e.message}`);
      }
    }

    // 将项目名称列表转换为字符串，每个项目一行
    const projectNamesText = projectNames.join("\n");
    ws[`F${rowNum}`] = { v: projectNamesText, t: "s" };
  }

  // 设置工作表范围
  const range = { s: { c: 0, r: 0 }, e: { c: 5, r: data.length + 1 } };
  ws["!ref"] = XLSX.utils.encode_range(range);

  // 设置列宽
  ws["!cols"] = [
    { wch: 12 }, // 姓名
    { wch: 8 }, // 年龄
    { wch: 18 }, // 身份证号
    { wch: 25 }, // 业绩
    { wch: 50 }, // 专业
    { wch: 35 }, // 项目名称
  ];

  // 设置行高
  const rowHeights = [];
  for (let i = 0; i <= data.length; i++) {
    rowHeights.push({ hpt: i === 0 ? 30 : 100 }); // 标题30pt，数据行100pt
  }
  ws["!rows"] = rowHeights;

  console.log(`Excel 工作表生成完成！`);
  return ws;
}
