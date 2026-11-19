// HackPropagation. Standalone Excel Office Script.
function main(workbook: ExcelScript.Workbook) {
    let _workbook: ExcelScript.Workbook = workbook;
    const designSheet = workbook.getWorksheet("Architect");
    if (!designSheet || !workbook.getWorksheet("Helper") || !workbook.getWorksheet("tinyDigits")) {
        throw new Error("Open the HackPropagation workbook included with this project before running this script.");
    }
    const epochCount = Number(designSheet.getRange("H5").getValue());
    if (!Number.isInteger(epochCount) || epochCount < 1 || epochCount > 200) {
        throw new Error("H5 must contain a whole epoch count from 1 to 200.");
    }
    const requestedLayers = designSheet.getRange("B5:B50").getValues().map(row => String(row[0]).trim());
    if (!["CELoss", "MSELoss"].includes(requestedLayers.filter(layer => layer !== "").pop() || "")) {
        throw new Error("HackPropagation requires CELoss or MSELoss as the final layer.");
    }
    designSheet.getRange("G5").setControl({ type: ExcelScript.CellControlType.empty });
    designSheet.getRange("G4").setValue("Mode");
    designSheet.getRange("G5").setValue("Backpropagation");
    designSheet.getRange("G:G").getFormat().setColumnWidth(120);
    designSheet.getRange("G6:K30").clear(ExcelScript.ClearApplyTo.contents);
    const sheets = workbook.getWorksheets();
    for (let i = sheets.length - 1; i >= 0; i--) {
        const sheet = sheets[i];
        const name = sheet.getName();
        if (/^Blueprint_ep\d+$/.test(name) || name === "Blueprint_inference") {
            sheet.delete();
        }
    }
    const helperSheet = workbook.getWorksheet("Helper");
    const range = helperSheet.getRange("AS2:AU50");
    range.clear(ExcelScript.ClearApplyTo.contents);
    let dashboardSheet = workbook.getWorksheet("Dashboard");
    if (dashboardSheet) {
        dashboardSheet.delete();
    }
    const source = designSheet.getRange("B5:D50").getValues();
    const outInputCol = 44;
    const outOpsCol = 45;
    const outOutputCol = 46;
    let dim0 = "";
    let batch = "";
    let inputPrefix = "";
    let lastLinearSourceIdx = -1;
    for (let j = 0; j < source.length; j++) {
        if (String(source[j][0] ?? "").trim() === "Linear") {
            lastLinearSourceIdx = j;
        }
    }
    const isTinyDigitsConv = String(source[0][0] ?? "").trim() === "tinyDigits";
    let outRow = 1;
    let firstOp = true;
    let layerInformation: {
        [key: number]: [
            string,
            number,
            number
        ];
    } = {};
    for (let i = 0; i < source.length; i++) {
        const layer = String(source[i][0] ?? "").trim();
        let dim = source[i][1];
        const batchVal = source[i][2];
        if (layer === "")
            continue;
        let _dim = Number(dim);
        if (i == 0) {
            if (layer === "Image") {
                layerInformation[i] = [layer, _dim * _dim, Number(source[0][2])];
            }
            else if (layer === "tinyDigits") {
                layerInformation[i] = [layer, 25, Number(source[0][2])];
            }
            else
                layerInformation[i] = [layer, _dim, Number(source[0][2])];
        }
        else
            layerInformation[i] = [layer, _dim > 0 ? _dim : layerInformation[i - 1][1], Number(source[0][2])];
        if (layer === "Rand") {
            dim0 = String(dim);
            batch = String(batchVal);
            inputPrefix = `X=Rand, feat=${dim0}, batch=${batch}`;
            continue;
        }
        if (layer === "Image") {
            dim0 = String(dim);
            batch = String(batchVal);
            inputPrefix = `X=Image, size=${dim0}, batch=${batch}`;
            continue;
        }
        if (layer === "tinyDigits") {
            batch = String(batchVal);
            inputPrefix = `X=tinyDigits, batch=${batch}`;
            designSheet.getCell(4 + i, 2).setValue("");
            continue;
        }
        let inputText = "";
        let outputText = "X";
        if (firstOp) {
            inputText = inputPrefix;
            firstOp = false;
        }
        else {
            inputText = "X";
        }
        if (layer === "Linear") {
            if (isTinyDigitsConv && i === lastLinearSourceIdx) {
                designSheet.getCell(4 + i, 2).setValue(10);
                dim = 10;
            }
            outputText = `X, dim=${dim}`;
        }
        if (layer === "CELoss") {
            inputText = "X";
            outputText = "L";
        }
        helperSheet.getCell(outRow, outInputCol).setValue(inputText);
        helperSheet.getCell(outRow, outOpsCol).setValue(layer);
        helperSheet.getCell(outRow, outOutputCol).setValue(outputText);
        outRow++;
    }
    const statusRange = designSheet.getRange("E5:F30");
    const values = range.getValues();
    const allInputs = values.map(row => row[0]);
    let allOperations = values.map(row => row[1]);
    allOperations = values
        .filter(row => row.every(cell => cell !== "" && cell !== null))
        .map(row => row[1]);
    const allOutputs = values.map(row => row[2]);
    designSheet.setTabColor("Yellow");
    designSheet.setPosition(0);
    designSheet.getCell(0, 4).setValue("");
    statusRange.clear(ExcelScript.ClearApplyTo.all);
    workbook.getNames().forEach(n => n.delete());
    workbook.getWorksheets().forEach(sheet => {
        sheet.getNames().forEach(n => n.delete());
    });
    let abhSheet = workbook.getWorksheet("Blueprint");
    if (abhSheet) {
        abhSheet.getUsedRange()?.clear();
        workbook.getNames().forEach(namedItem => {
            const formula = namedItem.getFormula();
            if (formula && formula.includes(`Blueprint`)) {
                namedItem.delete();
            }
        });
        workbook.getNames().forEach(n => {
            try {
                const ref = n.getRange();
            }
            catch {
                n.delete();
            }
        });
    }
    else {
        abhSheet = workbook.addWorksheet("Blueprint");
    }
    const backpropFlag = true;
    let numOfEpoch = designSheet.getRange("H5").getValue();
    numOfEpoch = String(Number(numOfEpoch) - 1);
    let totalJobs = backpropFlag == true ? Number(numOfEpoch) : -1;
    desighedText(designSheet, "🔁 Working.....", 10 + totalJobs, 6, 1, 1, "white", "Ink Free", 11, "gray", false, "left");
    abhSheet.setTabColor("RoyalBlue");
    abhSheet.setPosition(1);
    let W = 20;
    let H = W;
    let shapeLeftRightPos = W;
    let shapeBottomTopPos = H;
    let shapeWidth = 50;
    let shapeHeight = 20;
    abhSheet.setShowGridlines(false);
    abhSheet.setShowHeadings(false);
    abhSheet.getRange().getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    abhSheet.getRange().getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    abhSheet.getRange().getFormat().setColumnWidth(W);
    abhSheet.getRange().getFormat().setRowHeight(H);
    abhSheet.getRange().getFormat().getFont().setName("Ink Free");
    abhSheet.getRange().getFormat().getFont().setSize(7);
    abhSheet.getRange().getFormat().getFont().setBold(false);
    abhSheet.getRange("A:A").getFormat().setColumnWidth(25);
    abhSheet.getRange("2:2").getFormat().setRowHeight(25);
    designSheet.getRange("A:A").getFormat().setColumnWidth(45);
    designSheet.getRange("2:2").getFormat().setRowHeight(25);
    let maxDim = 1;
    let maxDimTrimmed = 1;
    titleMaker(abhSheet, "HackPropagation · Forward + Backprop", 0, 2);
    titleMaker(designSheet, "✍️ HackPropagation", 0, 1);
    const colStart = 2;
    const rowStart = 3;
    let cursor = { row: rowStart, col: colStart, range };
    let blockStartCursor = { row: rowStart, col: colStart, range };
    let blockEndCursor = { row: rowStart, col: colStart, range };
    let T = 0;
    let d = 0;
    let lrrow = rowStart;
    let lrcol = colStart;
    let hasInput: boolean = false;
    let namedVariable = "X";
    let output_var_name = "X";
    let dim_var_name = "d";
    let dim_value = 1;
    let temprange: ExcelScript.Range;
    let currentShape: ExcelScript.Shape;
    let prevShape: ExcelScript.Shape;
    let params: {
        [key: string]: [
            string,
            number
        ];
    } = {};
    const regex = /^(\w+),?\s*dim\s+(\w+)\s*=\s*(\d+)/i;
    for (let i = 0; i < allOutputs.length; i++) {
        const cellValue = String(allOutputs[i]).trim();
        const match = cellValue.match(regex);
        if (match) {
            const outVarName = match[1];
            const key = match[2];
            const num = Number(match[3]);
            params[outVarName] = [key, num];
        }
        else {
            const simpleMatch = cellValue.match(/^(\w+)$/);
            if (simpleMatch) {
                const outVarName = simpleMatch[1];
                if (!(outVarName in params)) {
                    params[outVarName] = ["", 1];
                }
            }
        }
    }
    let numberOfClasses = 0;
    let samplesPerClasses = 0;
    let imRow = 0;
    let imCol = 0;
    let ops_var_temp = "";
    let torchCode = "";
    const showCode = false;
    let showList = false;
    let dictVar: {
        [key: string]: string;
    } = {};
    let allCode: string[] = [];
    allCode.push("import torch");
    allCode.push("import torch.nn as nn");
    allCode.push("");
    const seqLayers: string[] = [];
    const seqLayerLinks: string[] = [];
    const seqLayerAnchors: string[][] = [];
    let lossFnAnchorNames: string[] = ["lossfn"];
    let lossEvalAnchorNames: string[] = [];
    let codeTopRow = 0;
    let inputBatch = 0;
    let inputFeatures = 0;
    let inputTensorCode = "";
    let nClasses = 0;
    let modelOutputLinkName = "";
    let lossType = "";
    let lastLinearOut = 0;
    let flattenedSize = 0;
    let codeShiftRight = 4;
    let generatedCodeStartCol = -1;
    let generatedCodeColCount = 0;
    let weightUpdateRightCol = -1;
    let backpropBlockSpans: {
        startRow: number;
        endRow: number;
        startCol: number;
        color: string;
    }[] = [];
    let lossVarName = "L";
    let endOfModel = 0;
    let forwardPass: {
        [key: number]: [
            string,
            string,
            string,
            string,
            number,
            number,
            number,
            number
        ];
    } = {};
    let forwardPassAttention: {
        [key: string]: [
            string,
            string,
            string,
            string,
            number,
            number,
            number,
            number
        ];
    } = {};
    let weightInformation: {
        [key: string]: [
            number,
            number,
            number,
            number,
            number
        ];
    } = {};
    let hasLoss = allOperations.some(op => typeof op === "string" && op.toLowerCase().includes("loss"));
    let number_of_ops = 0;
    let fpass_counter = 0;
    let numOfLayer = values.length;
    for (let i = 0; i < values.length; i++) {
        let localCode: string[] = [];
        const row = values[i];
        const input = row[0] as string;
        const operation = row[1] as string;
        const output = row[2] as string;
        let input_var_name = "";
        let input_string = "";
        let out_dim = 1;
        let out_dimk = 1;
        const regex_dim = /^(\w+),?\s*dim\s+(\w+)\s*=\s*(\d+)/i;
        const regex_dim2 = /^(\w+),\s*dim\s*=\s*(\d+)$/i;
        const regex_dim3 = /^(\w+),\s*dims\s*\(\s*((?:\w+\s*=\s*\d+\s*,\s*)*\w+\s*=\s*\d+)\s*\)\s*$/i;
        const match = output.match(regex_dim);
        const match2 = output.match(regex_dim2);
        const match3 = output.match(regex_dim3);
        if (match) {
            out_dim = Number(match[3]);
        }
        else if (match2) {
            out_dim = Number(match2[2]);
        }
        else if (match3) {
            const varName = match3[1];
            const inside = match3[2];
            const dims: {
                [key: string]: number;
            } = {};
            inside.split(",").forEach(pair => {
                const [k, v] = pair.split("=");
                dims[k.trim().toLowerCase()] = Number(v.trim());
            });
            out_dimk = dims["k"];
            out_dim = dims["v"];
        }
        if (input == "") {
            break;
        }
        if (input) {
            const input_token_match = input.match(/^([^=]+)=([^=]+)$/);
            if (input_token_match) {
                const xd = 1;
                input_var_name = input_token_match[1].trim();
                input_string = input_token_match[2].trim();
                T = input_string.length;
                namedVariable = input_var_name;
                cursor = drawXwithText(abhSheet, input_var_name, input_string, cursor.row, cursor.col, xd, T);
                updateName(workbook, cursor, output_var_name, abhSheet);
                let range_: ExcelScript.Range;
                range_ = cursor.range;
                let tcode = getCode(range_, input_var_name, output_var_name, cursor.row, cursor.col, xd, T);
                const cc = cursor.col + (T + 2);
                const cr = cursor.row - (xd + 1);
                abhSheet.getCell(cr, cc).setValue(tcode);
                abhSheet.getCell(cr, cc).getFormat().getFont().setName("Courier New");
                abhSheet.getCell(cr, cc).getFormat().getFont().setColor("gray");
                abhSheet.getCell(cr, cc).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
            }
            else {
                input_var_name = input;
            }
            const input_token_match_im = input.match(/^\s*(\w+)\s*=\s*(\w+)\s*,\s*batch\s*=\s*(\d+)/i);
            if (input_token_match_im) {
                input_var_name = input_token_match_im[1];
                const datasetName = input_token_match_im[2];
                const batchNumber = Number(input_token_match_im[3]);
                T = batchNumber;
                numberOfClasses = 10;
                samplesPerClasses = 11;
                imRow = 5;
                imCol = 5;
                const regex_2 = /^(\w+),\s*dim\s*=\s*(\d+)$/i;
                const nums: number[] = [];
                nums.push(imRow * imCol);
                for (let i = 0; i < allOutputs.length; i++) {
                    const cellValue = String(allOutputs[i]).trim();
                    const match = cellValue.match(regex_2);
                    if (match) {
                        const num = Number(match[2]);
                        nums.push(num);
                    }
                }
                maxDim = Math.max(...nums);
                maxDimTrimmed = Math.max(...nums.slice(1));
                cursor.col = maxDim + 5;
                const lastDim = nums[nums.length - 1];
                const next_ops = values[i][1] as string;
                const includeBias = next_ops === "Linear";
                cursor = displayImages2(workbook, abhSheet, `${input_var_name} = ${datasetName} (${imRow}x${imCol})`, cursor.row, cursor.col, datasetName, numberOfClasses, samplesPerClasses, batchNumber, imRow, imCol, includeBias, hasLoss, maxDim);
                localCode.push("import torch");
                localCode.push("import torch.nn as nn");
                torchCode = genTorchCode("Image", output_var_name, String(T), String(imRow));
                allCode.push(torchCode);
                localCode.push(torchCode);
                if (hasLoss) {
                    torchCode = genTorchCode("oheImage", "target", String(numberOfClasses), String(T));
                    allCode.push(torchCode);
                    localCode.push(torchCode);
                }
                if (showCode) {
                    inputBatch = T;
                    inputFeatures = imRow * imCol;
                    inputTensorCode = `X = torch.randn(${T}, 1, ${imRow}, ${imCol})`;
                    nClasses = Number(numberOfClasses);
                    codeTopRow = cursor.row - 3;
                }
                updateName(workbook, cursor, input_var_name, abhSheet);
            }
            const input_token_match_imrand = input.match(/^\s*(\w+)\s*=\s*image\s*,\s*size\s*=\s*(\d+)\s*,\s*batch\s*=\s*(\d+)\s*$/i);
            if (input_token_match_imrand) {
                input_var_name = input_token_match_imrand[1];
                const randImageSize = Number(input_token_match_imrand[2]);
                const batchNumber = Number(input_token_match_imrand[3]);
                T = batchNumber;
                imRow = randImageSize;
                imCol = randImageSize;
                const regex_2 = /^(\w+),\s*dim\s*=\s*(\d+)$/i;
                const nums: number[] = [];
                nums.push(imRow * imCol);
                for (let i = 0; i < allOutputs.length; i++) {
                    const cellValue = String(allOutputs[i]).trim();
                    const match = cellValue.match(regex_2);
                    if (match) {
                        const num = Number(match[2]);
                        nums.push(num);
                    }
                }
                maxDim = Math.max(...nums);
                maxDimTrimmed = Math.max(...nums.slice(1));
                cursor.col = maxDim + 5;
                const lastDim = nums[nums.length - 1];
                const next_ops = values[i][1] as string;
                const includeBias = next_ops === "Linear";
                cursor = displayRandImages2(workbook, abhSheet, `${input_var_name} = Image (${imRow}x${imCol})`, cursor.row, cursor.col, "", lastDim, samplesPerClasses, T, imRow, imCol, includeBias, hasLoss, maxDim, "seeded", 101);
                localCode.push("import torch");
                localCode.push("import torch.nn as nn");
                torchCode = genTorchCode("Image", output_var_name, String(T), String(imRow));
                allCode.push(torchCode);
                localCode.push(torchCode);
                if (hasLoss) {
                    torchCode = genTorchCode("oheImage", "target", String(numberOfClasses), String(T));
                    allCode.push(torchCode);
                    localCode.push(torchCode);
                }
                if (showCode) {
                    inputBatch = T;
                    inputFeatures = imRow * imCol;
                    inputTensorCode = `X = torch.randn(${T}, 1, ${imRow}, ${imCol})`;
                    nClasses = Number(numberOfClasses);
                    codeTopRow = cursor.row - 3;
                }
                updateName(workbook, cursor, input_var_name, abhSheet);
            }
            const input_token_match_rand = input.match(/^\s*(\w+)\s*=\s*Rand\s*,\s*feat\s*=\s*(\d+)\s*,\s*batch\s*=\s*(\d+)/i);
            if (input_token_match_rand) {
                input_var_name = input_token_match_rand[1];
                const features = Number(input_token_match_rand[2]);
                const batchNumber = Number(input_token_match_rand[3]);
                T = batchNumber;
                const regex_2 = /^(\w+),\s*dim\s*=\s*(\d+)$/i;
                const regex_3 = /^(\w+),\s*dims\s*\(\s*((?:\w+\s*=\s*\d+\s*,\s*)*\w+\s*=\s*\d+)\s*\)\s*$/i;
                const nums: number[] = [];
                nums.push(features);
                for (let i = 0; i < allOutputs.length; i++) {
                    const cellValue = String(allOutputs[i]).trim();
                    const match_3 = cellValue.match(regex_3);
                    if (match_3) {
                        const inside = match_3[2];
                        const dims: {
                            [key: string]: number;
                        } = {};
                        inside.split(",").forEach(pair => {
                            const [k, v] = pair.split("=");
                            dims[k.trim().toLowerCase()] = Number(v.trim());
                        });
                        nums.push(dims["k"]);
                        nums.push(T);
                        nums.push(dims["v"]);
                    }
                    const match = cellValue.match(regex_2);
                    if (match) {
                        const num = Number(match[2]);
                        nums.push(num);
                    }
                }
                const lastDim = nums[nums.length - 1];
                maxDim = Math.max(...nums);
                maxDimTrimmed = Math.max(...nums.slice(1));
                cursor.col = maxDim + 5;
                const next_ops = values[i][1] as string;
                const includeBias = next_ops === "Linear";
                cursor = RandInput2(workbook, abhSheet, input_var_name + "=Rand", cursor.row, cursor.col, features, T, lastDim, includeBias, "seeded", hasLoss, 102);
                localCode.push("import torch");
                localCode.push("import torch.nn as nn");
                torchCode = genTorchCode("Rand", output_var_name, String(T), String(features));
                allCode.push(torchCode);
                localCode.push(torchCode);
                let code_up = 1;
                if (hasLoss) {
                    torchCode = genTorchCode("ohe", "target", String(lastDim), String(T));
                    allCode.push(torchCode);
                    localCode.push(torchCode);
                    allCode.push(torchCode);
                    code_up = lastDim + 1;
                }
                if (showCode) {
                    inputBatch = Number(T);
                    inputFeatures = Number(features);
                    inputTensorCode = `X = torch.randn(${T}, ${features})`;
                    if (hasLoss)
                        nClasses = Number(lastDim);
                    codeTopRow = cursor.row - code_up - 2;
                }
                updateName(workbook, cursor, input_var_name, abhSheet);
            }
        }
        if (output) {
            const output_match = output.match(/^(\w+),\s*dim\s+(\w+)\s*=\s*(\d+)$/i);
            const output_match2 = output.match(/^(\w+),\s*dim\s*=\s*(\d+)$/i);
            const output_match3 = output.match(/^(\w+),\s*dims\s*\(\s*((?:\w+\s*=\s*\d+\s*,\s*)*\w+\s*=\s*\d+)\s*\)\s*$/i);
            if (output_match) {
                output_var_name = output_match[1];
                dim_var_name = output_match[2];
                dim_value = Number(output_match[3]);
                namedVariable = output_var_name;
            }
            else if (output_match2) {
                output_var_name = output_match2[1];
                dim_value = Number(output_match2[2]);
                namedVariable = output_var_name;
            }
            else if (output_match3) {
                output_var_name = output_match3[1];
                namedVariable = output_var_name;
            }
            else {
                output_var_name = output;
                namedVariable = output_var_name;
            }
            let d1 = out_dim;
            blockStartCursor = cursor;
            var codeLinkTarget = getLayerCodeAnchor(i);
            var codeLinkLabel = "</>";
            cursor.row += 1;
            if (number_of_ops == 0) {
                localCode = [];
                allCode.push("model=nn.Sequential(");
            }
            if (operation == "Flatten") {
                if (number_of_ops == 0) {
                    output_var_name = output_var_name + "_" + String(i);
                }
                else {
                    input_var_name = input_var_name + "_" + String(i - 1);
                    output_var_name = output_var_name + "_" + String(i);
                }
                const input_var_name_sub = toSubscript(input_var_name);
                const output_var_name_sub = toSubscript(output_var_name);
                let r = abhSheet.getNamedItem(input_var_name).getRange();
                if (imRow == 0 && imCol == 0) {
                    imRow = r.getRowCount();
                    imCol = 1;
                }
                forwardPass[fpass_counter] = [
                    input_var_name,
                    operation,
                    String(values[i + 1][1]),
                    output_var_name,
                    cursor.row, cursor.col,
                    imRow * imCol, T
                ];
                fpass_counter++;
                if (i + 1 < values.length) {
                    const next_ops = values[i + 1][1] as string;
                    if (next_ops == "Linear") {
                        cursor = Flattening(abhSheet, changeTitle2(output_var_name_sub, operation, input_var_name_sub), cursor.row, cursor.col, r, imRow, imCol, T, true, input_var_name);
                        updateName(workbook, cursor, output_var_name, abhSheet);
                    }
                    else {
                        cursor = Flattening(abhSheet, changeTitle2(output_var_name_sub, operation, input_var_name_sub), cursor.row, cursor.col, r, imRow, imCol, T, false, input_var_name);
                        updateName(workbook, cursor, output_var_name, abhSheet);
                    }
                }
                const coderVar = "L" + number_of_ops + "_" + operation;
                torchCode = genTorchCode(operation, operation.toLowerCase() + String(i), String(T), "");
                allCode.push("    " + torchCode + ",");
                localCode.push(torchCode);
                const fcode = genzForwardPassCode(operation.toLowerCase() + String(i), input_var_name, output_var_name);
                localCode.push(fcode);
                if (showCode) {
                    seqLayers.push("    " + genTorchCode(operation));
                    seqLayerLinks.push(output_var_name);
                    seqLayerAnchors.push([getLayerCodeAnchor(i)]);
                    modelOutputLinkName = output_var_name;
                    modelOutputLinkName = output_var_name;
                    flattenedSize = imRow * imCol;
                }
                number_of_ops++;
            }
            else if (operation == "LayerNorm") {
                if (number_of_ops == 0) {
                    output_var_name = output_var_name + "_" + String(i);
                }
                else {
                    input_var_name = input_var_name + "_" + String(i - 1);
                    output_var_name = output_var_name + "_" + String(i);
                }
                const input_var_name_sub = toSubscript(input_var_name);
                const output_var_name_sub = toSubscript(output_var_name);
                const commCell = abhSheet.getCell(cursor.row - 2, cursor.col - 1);
                const r = abhSheet.getNamedItem(input_var_name).getRange();
                const rows = r.getRowCount();
                T = r.getColumnCount();
                forwardPass[fpass_counter] = [
                    input_var_name,
                    operation,
                    String(values[i + 1][1]),
                    output_var_name,
                    cursor.row, cursor.col,
                    rows, T
                ];
                fpass_counter++;
                const next_ops = values[i + 1][1] as string;
                let header = changeTitle2(output_var_name_sub, operation, input_var_name_sub);
                const includeBias = next_ops === "Linear";
                cursor = LayerNorm(abhSheet, header, cursor.row, cursor.col, r, rows, T, includeBias, input_var_name_sub, i);
                updateName(workbook, cursor, output_var_name, abhSheet);
                let temp_row = cursor.row - rows - 2 - 1 * Number(includeBias);
                let temp_col = cursor.col;
                const range_std = abhSheet.getRange(convertToA1(temp_row, temp_col) + ":" + convertToA1(temp_row, temp_col + T - 1));
                workbook.addNamedItem("std_" + input_var_name, range_std);
                temp_row -= 2;
                const range_mean = abhSheet.getRange(convertToA1(temp_row, temp_col) + ":" + convertToA1(temp_row, temp_col + T - 1));
                workbook.addNamedItem("mean_" + input_var_name, range_mean);
                const coderVar = "L" + number_of_ops + "_" + operation;
                torchCode = genTorchCode(operation, operation.toLowerCase() + String(i), String(rows), "");
                allCode.push("    " + torchCode + ",");
                localCode.push(torchCode);
                const fcode = genzForwardPassCode(operation.toLowerCase() + String(i), input_var_name, output_var_name);
                localCode.push(fcode);
                if (showCode) {
                    seqLayers.push("    " + genTorchCode(operation, "", String(rows)));
                    seqLayerLinks.push(output_var_name);
                    seqLayerAnchors.push([getLayerCodeAnchor(i)]);
                    modelOutputLinkName = output_var_name;
                }
                number_of_ops++;
            }
            else if (operation == "⟶") {
                ;
            }
            else if (operation == "ReLU") {
                if (number_of_ops == 0) {
                    output_var_name = output_var_name + "_" + String(i);
                }
                else {
                    input_var_name = input_var_name + "_" + String(i - 1);
                    output_var_name = output_var_name + "_" + String(i);
                }
                const input_var_name_sub = toSubscript(input_var_name);
                const output_var_name_sub = toSubscript(output_var_name);
                const commCell = abhSheet.getCell(cursor.row - 2, cursor.col - 1);
                const r = abhSheet.getNamedItem(input_var_name).getRange();
                const rows = r.getRowCount();
                T = r.getColumnCount();
                forwardPass[fpass_counter] = [
                    input_var_name,
                    operation,
                    String(values[i + 1][1]),
                    output_var_name,
                    cursor.row, cursor.col,
                    rows, T
                ];
                fpass_counter++;
                const next_ops = values[i + 1][1] as string;
                const includeBias = next_ops === "Linear";
                cursor = ReLU(abhSheet, changeTitle2(output_var_name_sub, operation, input_var_name_sub), cursor.row, cursor.col, r, rows, T, includeBias, input_var_name_sub);
                updateName(workbook, cursor, output_var_name, abhSheet);
                const coderVar = "L" + number_of_ops + "_" + operation;
                torchCode = genTorchCode(operation, operation.toLowerCase() + String(i), String(T), "");
                allCode.push("    " + torchCode + ",");
                localCode.push(torchCode);
                const fcode = genzForwardPassCode(operation.toLowerCase() + String(i), input_var_name, output_var_name);
                localCode.push(fcode);
                if (showCode) {
                    seqLayers.push("    " + genTorchCode(operation));
                    seqLayerLinks.push(output_var_name);
                    seqLayerAnchors.push([getLayerCodeAnchor(i)]);
                    modelOutputLinkName = output_var_name;
                }
                number_of_ops++;
            }
            else if (operation == "Tanh") {
                if (number_of_ops == 0) {
                    output_var_name = output_var_name + "_" + String(i);
                }
                else {
                    input_var_name = input_var_name + "_" + String(i - 1);
                    output_var_name = output_var_name + "_" + String(i);
                }
                const input_var_name_sub = toSubscript(input_var_name);
                const output_var_name_sub = toSubscript(output_var_name);
                const commCell = abhSheet.getCell(cursor.row - 2, cursor.col - 1);
                const r = abhSheet.getNamedItem(input_var_name).getRange();
                const rows = r.getRowCount();
                T = r.getColumnCount();
                forwardPass[fpass_counter] = [
                    input_var_name,
                    operation,
                    String(values[i + 1][1]),
                    output_var_name,
                    cursor.row, cursor.col,
                    rows, T
                ];
                fpass_counter++;
                const next_ops = values[i + 1][1] as string;
                const includeBias = next_ops === "Linear";
                cursor = TanH(abhSheet, changeTitle2(output_var_name_sub, operation, input_var_name_sub), cursor.row, cursor.col, r, rows, T, includeBias, input_var_name_sub);
                updateName(workbook, cursor, output_var_name, abhSheet);
                const coderVar = "L" + number_of_ops + "_" + operation;
                torchCode = genTorchCode(operation, operation.toLowerCase() + String(i), String(T), "");
                allCode.push("    " + torchCode + ",");
                localCode.push(torchCode);
                const fcode = genzForwardPassCode(operation.toLowerCase() + String(i), input_var_name, output_var_name);
                localCode.push(fcode);
                if (showCode) {
                    seqLayers.push("    " + genTorchCode(operation));
                    seqLayerLinks.push(output_var_name);
                    seqLayerAnchors.push([getLayerCodeAnchor(i)]);
                    modelOutputLinkName = output_var_name;
                }
                number_of_ops++;
            }
            else if (operation == "GELU") {
                if (number_of_ops == 0) {
                    output_var_name = output_var_name + "_" + String(i);
                }
                else {
                    input_var_name = input_var_name + "_" + String(i - 1);
                    output_var_name = output_var_name + "_" + String(i);
                }
                const input_var_name_sub = toSubscript(input_var_name);
                const output_var_name_sub = toSubscript(output_var_name);
                const commCell = abhSheet.getCell(cursor.row - 2, cursor.col - 1);
                const r = abhSheet.getNamedItem(input_var_name).getRange();
                const rows = r.getRowCount();
                T = r.getColumnCount();
                forwardPass[fpass_counter] = [
                    input_var_name,
                    operation,
                    String(values[i + 1][1]),
                    output_var_name,
                    cursor.row, cursor.col,
                    rows, T
                ];
                fpass_counter++;
                const next_ops = values[i + 1][1] as string;
                const includeBias = next_ops === "Linear";
                const temprow = cursor.row + 2;
                const tempcol = cursor.col;
                cursor = GELUerf(abhSheet, changeTitle2(output_var_name_sub, operation, input_var_name_sub), cursor.row, cursor.col, r, rows, T, includeBias, input_var_name_sub, i);
                updateName(workbook, cursor, output_var_name, abhSheet);
                const erf_range = abhSheet.getRangeByIndexes(temprow, tempcol, rows, T);
                const tempcur = { row: temprow, col: tempcol, range: erf_range };
                updateName(workbook, tempcur, `erf_${i}`, abhSheet);
                const coderVar = "L" + number_of_ops + "_" + operation;
                torchCode = genTorchCode(operation, operation.toLowerCase() + String(i), String(T), "");
                allCode.push("    " + torchCode + ",");
                localCode.push(torchCode);
                const fcode = genzForwardPassCode(operation.toLowerCase() + String(i), input_var_name, output_var_name);
                localCode.push(fcode);
                if (showCode) {
                    seqLayers.push("    " + genTorchCode(operation));
                    seqLayerLinks.push(output_var_name);
                    seqLayerAnchors.push([getLayerCodeAnchor(i)]);
                    modelOutputLinkName = output_var_name;
                }
                number_of_ops++;
            }
            else if (operation == "Softmax") {
                if (number_of_ops == 0) {
                    output_var_name = output_var_name + "_" + String(i);
                }
                else {
                    input_var_name = input_var_name + "_" + String(i - 1);
                    output_var_name = output_var_name + "_" + String(i);
                }
                const input_var_name_sub = toSubscript(input_var_name);
                const output_var_name_sub = toSubscript(output_var_name);
                const commCell = abhSheet.getCell(cursor.row - 2, cursor.col - 1);
                const r = abhSheet.getNamedItem(input_var_name).getRange();
                const rows = r.getRowCount();
                T = r.getColumnCount();
                forwardPass[fpass_counter] = [
                    input_var_name,
                    operation,
                    String(values[i + 1][1]),
                    output_var_name,
                    cursor.row, cursor.col,
                    rows, T
                ];
                fpass_counter++;
                const next_ops = values[i + 1][1] as string;
                const includeBias = next_ops === "Linear";
                cursor = softmax(abhSheet, changeTitle2(output_var_name_sub, operation, input_var_name_sub), cursor.row, cursor.col, r, rows, T, includeBias, input_var_name_sub);
                updateName(workbook, cursor, output_var_name, abhSheet);
                const coderVar = "L" + number_of_ops + "_" + operation;
                torchCode = genTorchCode(operation, operation.toLowerCase() + String(i), String(T), "");
                if (next_ops != "CELoss") {
                    allCode.push("    " + torchCode + ",");
                    localCode.push(torchCode);
                    const fcode = genzForwardPassCode(operation.toLowerCase() + String(i), input_var_name, output_var_name);
                    localCode.push(fcode);
                    if (showCode) {
                        seqLayers.push("    " + genTorchCode(operation));
                        seqLayerLinks.push(output_var_name);
                        seqLayerAnchors.push([getLayerCodeAnchor(i)]);
                        modelOutputLinkName = output_var_name;
                    }
                }
                else if (next_ops == "CELoss") {
                    let localCode: string[] = [];
                    torchCode = genTorchCode("CELoss");
                    allCode.push(torchCode);
                    localCode.push(torchCode);
                    torchCode = genTorchCode("loss", String(values[i + 1][2]), input_var_name, "target");
                    allCode.push(torchCode);
                    localCode.push(torchCode);
                    lossVarName = String(values[i + 1][2]);
                    if (showCode) {
                        lossType = "CELoss";
                        modelOutputLinkName = output_var_name;
                        lossFnAnchorNames = ["lossfn", getLayerCodeAnchor(i)];
                        codeLinkLabel = "</>";
                    }
                }
                number_of_ops++;
            }
            else if (operation == "CELoss") {
                if (number_of_ops == 0) {
                    output_var_name = output_var_name + "_" + String(i);
                }
                else {
                    input_var_name = input_var_name + "_" + String(i - 1);
                    output_var_name = output_var_name + "_" + String(i);
                }
                const input_var_name_sub = toSubscript(input_var_name);
                const output_var_name_sub = toSubscript(output_var_name);
                const commCell = abhSheet.getCell(cursor.row - 2, cursor.col - 1);
                const r = abhSheet.getNamedItem(input_var_name).getRange();
                const r_ohe = abhSheet.getNamedItem("ohe").getRange();
                const rows = r.getRowCount();
                T = r.getColumnCount();
                const next_ops = values[i + 1][1] as string;
                const includeBias = next_ops === "Linear";
                cursor = CELoss(abhSheet, changeTitle2(output_var_name_sub, operation, input_var_name_sub), cursor.row, cursor.col, r, r_ohe, rows, T, includeBias, input_var_name_sub);
                updateName(workbook, cursor, output_var_name, abhSheet);
                if (values[i - 1][1] != "Softmax") {
                    let localCode: string[] = [];
                    torchCode = genTorchCode(operation);
                    allCode.push(torchCode);
                    localCode.push(torchCode);
                    torchCode = genTorchCode("loss", output_var_name, input_var_name, "target");
                    allCode.push(torchCode);
                    localCode.push(torchCode);
                    lossVarName = output_var_name;
                    if (showCode) {
                        lossType = operation;
                        modelOutputLinkName = input_var_name;
                        lossEvalAnchorNames = [getLayerCodeAnchor(i)];
                    }
                }
                else if (showCode) {
                    lossVarName = output_var_name;
                    lossFnAnchorNames = [...lossFnAnchorNames, getLayerCodeAnchor(i)];
                }
                number_of_ops++;
            }
            else if (operation == "MSELoss") {
                if (number_of_ops == 0) {
                    output_var_name = output_var_name + "_" + String(i);
                }
                else {
                    input_var_name = input_var_name + "_" + String(i - 1);
                    output_var_name = output_var_name + "_" + String(i);
                }
                const input_var_name_sub = toSubscript(input_var_name);
                const output_var_name_sub = toSubscript(output_var_name);
                const commCell = abhSheet.getCell(cursor.row - 2, cursor.col - 1);
                const r = abhSheet.getNamedItem(input_var_name).getRange();
                const r_ohe = abhSheet.getNamedItem("ohe").getRange();
                const rows = r.getRowCount();
                T = r.getColumnCount();
                const next_ops = values[i + 1][1] as string;
                const includeBias = next_ops === "Linear";
                cursor = MSELoss(abhSheet, changeTitle2(output_var_name_sub, operation, input_var_name_sub), cursor.row, cursor.col, r, r_ohe, rows, T, includeBias, input_var_name_sub);
                updateName(workbook, cursor, output_var_name, abhSheet);
                let localCode: string[] = [];
                torchCode = genTorchCode(operation);
                allCode.push(torchCode);
                localCode.push(torchCode);
                torchCode = genTorchCode("loss", output_var_name, input_var_name, "target");
                allCode.push(torchCode);
                localCode.push(torchCode);
                lossVarName = output_var_name;
                if (showCode) {
                    lossType = operation;
                    modelOutputLinkName = input_var_name;
                    lossEvalAnchorNames = [getLayerCodeAnchor(i)];
                }
                number_of_ops++;
            }
            else if (operation == "Linear") {
                if (number_of_ops == 0) {
                    output_var_name = output_var_name + "_" + String(i);
                }
                else {
                    input_var_name = input_var_name + "_" + String(i - 1);
                    output_var_name = output_var_name + "_" + String(i);
                }
                const input_var_name_sub = toSubscript(input_var_name);
                const output_var_name_sub = toSubscript(output_var_name);
                const next_ops = values[i + 1][1] as string;
                const includeBias = next_ops === "Linear";
                const isLastLinear = !values.slice(i + 1).some(r => r[1] === "Linear");
                const isTinyDigits = String(allInputs[0]).includes("tinyDigits");
                const d2 = (isTinyDigits && isLastLinear) ? 10 : out_dim;
                const range1 = abhSheet.getNamedItem(input_var_name).getRange();
                d1 = range1.getRowCount();
                T = range1.getColumnCount();
                forwardPass[fpass_counter] = [
                    input_var_name,
                    operation,
                    String(values[i + 1][1]),
                    output_var_name,
                    cursor.row, cursor.col,
                    d2, T
                ];
                fpass_counter++;
                cursor.col = cursor.col - (d1 + 1);
                abhSheet.getCell(cursor.row, cursor.col + d1 - 1)
                    .getFormat()
                    .getFont()
                    .setBold(false);
                abhSheet.getCell(cursor.row + 1, cursor.col + d1 - 1)
                    .getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.bottom);
                weightInformation["W" + "_" + String(i)] = [i, cursor.row + 2, cursor.col, d2, d1];
                cursor = drawMatrix3(abhSheet, toSubscript("W" + "_" + String(i)) + "|" + toSubscript("b" + "_" + String(i)), cursor.row, cursor.col, d2, d1, "seeded", undefined, 42, -1, 1);
                updateName(workbook, cursor, "W" + "_" + String(i), abhSheet);
                cursor.row = cursor.row - (d2 + 2);
                cursor.col = cursor.col + (d1 + 1);
                const range2 = abhSheet.getNamedItem("W" + "_" + String(i)).getRange();
                const linear_eq = toSubscript("W" + "_" + String(i)) + " " + input_var_name_sub + " + " + toSubscript("b" + "_" + String(i));
                makeArrow(abhSheet, "left2right", cursor.row + 1, cursor.col - 1, d2);
                cursor = multMatricesWithRand2(abhSheet, range2, range1, d1, d2, T, changeTitle2(output_var_name_sub, operation, input_var_name_sub), cursor.row, cursor.col, includeBias, linear_eq, input_var_name_sub);
                cursor.col = cursor.col - (d1 + 1);
                updateName(workbook, cursor, output_var_name, abhSheet);
                cursor.col = cursor.col + (d1 + 1);
                let localCode: string[] = [];
                const coderVar = "L" + number_of_ops + "_" + operation;
                torchCode = genTorchCode("Linear", operation.toLowerCase() + String(i), String(d1 - 1), String(d2));
                allCode.push("    " + torchCode + ",");
                localCode.push(torchCode);
                const fcode = genzForwardPassCode(operation.toLowerCase() + String(i), input_var_name, output_var_name);
                localCode.push(fcode);
                if (showCode) {
                    const linearIn = flattenedSize > 0 ? flattenedSize : Number(d1) - 1;
                    flattenedSize = 0;
                    seqLayers.push("    " + genTorchCode("Linear", "", String(linearIn), String(d2)));
                    seqLayerLinks.push(output_var_name);
                    seqLayerAnchors.push([getLayerCodeAnchor(i)]);
                    modelOutputLinkName = output_var_name;
                    lastLinearOut = Number(d2);
                }
                localCode.push(torchCode);
                number_of_ops++;
            }
            else if (operation == "PE") {
                if (number_of_ops == 0) {
                    output_var_name = output_var_name + "_" + String(i);
                }
                else {
                    input_var_name = input_var_name + "_" + String(i - 1);
                    output_var_name = output_var_name + "_" + String(i);
                }
                const input_var_name_sub = toSubscript(input_var_name);
                const output_var_name_sub = toSubscript(output_var_name);
                const next_ops = values[i + 1][1] as string;
                const includeBias = next_ops === "Linear";
                const range1 = abhSheet.getNamedItem(input_var_name).getRange();
                d1 = range1.getRowCount();
                T = range1.getColumnCount();
                cursor = PE(abhSheet, changeTitle2(output_var_name_sub, operation, input_var_name_sub), cursor.row, cursor.col, range1, d1, T, includeBias, input_var_name);
                updateName(workbook, cursor, output_var_name, abhSheet);
                number_of_ops++;
            }
            else if (operation == "Attention") {
                if (number_of_ops == 0) {
                    output_var_name = output_var_name + "_" + String(i);
                }
                else {
                    input_var_name = input_var_name + "_" + String(i - 1);
                    output_var_name = output_var_name + "_" + String(i);
                }
                const input_var_name_sub = toSubscript(input_var_name);
                const output_var_name_sub = toSubscript(output_var_name);
                const next_ops = values[i + 1][1] as string;
                const includeBias = next_ops === "Linear";
                let d2 = out_dim;
                const qkvnames: string[] = ["Q", "K", "V"];
                const dims: number[] = [out_dimk, out_dimk, d2];
                const range1 = workbook.getNamedItem(input_var_name).getRange();
                d1 = range1.getRowCount();
                T = range1.getColumnCount();
                for (let f = 0; f < qkvnames.length; f++) {
                    d2 = dims[f];
                    cursor.col = cursor.col - (d1 + 1);
                    abhSheet.getCell(cursor.row + 1, cursor.col + d1 - 1).setValue(toSubscript("b" + qkvnames[f].toLowerCase() + "_" + String(i)));
                    abhSheet.getCell(cursor.row, cursor.col + d1 - 1)
                        .getFormat()
                        .getFont()
                        .setBold(false);
                    abhSheet.getCell(cursor.row + 1, cursor.col + d1 - 1)
                        .getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.bottom);
                    forwardPass[fpass_counter] = [
                        input_var_name,
                        `${qkvnames[f]}`,
                        "AttDot",
                        `${qkvnames[f]}_${i}`,
                        cursor.row, cursor.col,
                        d2, T
                    ];
                    fpass_counter++;
                    cursor = drawMatrix3(abhSheet, toSubscript("W" + qkvnames[f].toLowerCase() + "_" + String(i)), cursor.row, cursor.col, d2, d1, "seeded", undefined, 42 - f, -1, 1);
                    updateName(workbook, cursor, "W" + qkvnames[f] + "_" + String(i));
                    cursor.row = cursor.row - (d2 + 2);
                    cursor.col = cursor.col + (d1 + 1);
                    const range2 = workbook.getNamedItem("W" + qkvnames[f] + "_" + String(i)).getRange();
                    const linear_eq = toSubscript("W" + qkvnames[f].toLowerCase() + "_" + String(i)) + " " + input_var_name_sub + " + " + toSubscript("b" + qkvnames[f].toLowerCase() + "_" + String(i));
                    cursor = multMatricesWithRand2(abhSheet, range2, range1, d1, d2, T, changeTitle2(toSubscript(qkvnames[f] + "_" + String(i)), "Linear", input_var_name_sub), cursor.row, cursor.col, false, linear_eq, input_var_name_sub);
                    cursor.col = cursor.col - (d1 + 1);
                    updateName(workbook, cursor, `${qkvnames[f]}_${i}`);
                    cursor.col = cursor.col + (d1 + 1);
                    cursor.row++;
                }
                const tempcursorrow = cursor.row;
                const tempcursorcol = cursor.col;
                const title_ = toSubscript(qkvnames[0] + "_" + String(i)) + "ᵀ" + toSubscript(qkvnames[1] + "_" + String(i));
                forwardPass[fpass_counter] = [
                    `${qkvnames[0]}_${i}` + ", " + `${qkvnames[1]}_${i}`,
                    "AttDot",
                    "AttDotScaled",
                    `AttDot_${i}`,
                    cursor.row, cursor.col,
                    d2, T
                ];
                fpass_counter++;
                cursor = multTwoMatrices2(abhSheet, toSubscript(`A_${i} = `), cursor.row, cursor.col, `${qkvnames[0]}_${i}`, true, `${qkvnames[1]}_${i}`, false, false, dims[0], T, dims[1], T);
                updateName(workbook, cursor, `AttDot_${i}`);
                cursor.row++;
                forwardPass[fpass_counter] = [
                    `AttDot_${i}`,
                    "AttDotScaled",
                    "Softmax",
                    `AttDotScaled_${i}`,
                    cursor.row, cursor.col,
                    d2, T
                ];
                fpass_counter++;
                cursor = scaleMatrix(abhSheet, toSubscript(`A_${i} = A_${i}/√${d1}`), cursor.row, cursor.col, Math.sqrt(d1), "div", `${convertToA1(cursor.row - T - 1, cursor.col)}#`, T, T);
                updateName(workbook, cursor, `AttDotScaled_${i}`);
                const r = workbook.getNamedItem(`AttDotScaled_${i}`).getRange();
                cursor.row++;
                forwardPass[fpass_counter] = [
                    `AttDotScaled_${i}`,
                    "Softmax",
                    "VA",
                    `AttWeight_${i}`,
                    cursor.row, cursor.col,
                    d2, T
                ];
                fpass_counter++;
                cursor = softmax(abhSheet, changeTitle2(toSubscript(`A_${i}`), `Softmax`, toSubscript(`A_${i}`)), cursor.row, cursor.col, r, T, T, false, toSubscript(`A_${i}`));
                updateName(workbook, cursor, `AttWeight_${i}`);
                cursor.row++;
                cursor = multTwoMatrices2(abhSheet, "", cursor.row, cursor.col, `${qkvnames[2]}_${i}`, false, `AttWeight_${i}`, false, false, dims[2], T, T, T, includeBias);
                updateName(workbook, cursor, output_var_name);
                forwardPass[fpass_counter] = [
                    `AttWeight_${i}`,
                    "VA",
                    String(values[i + 1][1]),
                    output_var_name,
                    cursor.row, cursor.col,
                    d2, T
                ];
                fpass_counter++;
                number_of_ops++;
            }
        }
        doneMessage(i + 4, designSheet, 3, false, output_var_name);
        blockEndCursor = cursor;
        if (showCode) goBlockLinks(abhSheet, cursor.row, cursor.col + T + 1, codeLinkTarget, codeLinkLabel);
        else goTop(abhSheet, cursor.row, cursor.col + T + 1);
        const verticalHeaderRange = abhSheet.getRange(convertToA1(blockStartCursor.row, 0) + ":" + convertToA1(blockEndCursor.row, 0));
        verticalHeaderRange.merge();
        verticalHeaderRange.getFormat().setTextOrientation(90);
        verticalHeaderRange.setValue(String(i) + ". " + operation);
        verticalHeaderRange.getFormat().getFont().setSize(16);
        verticalHeaderRange.getFormat().getFont().setName("Ink Free");
        verticalHeaderRange.getFormat().getFont().setColor("White");
        verticalHeaderRange.getFormat().getFont().setBold(true);
        verticalHeaderRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
        verticalHeaderRange.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
        const blockrange = abhSheet.getRange(convertToA1(blockStartCursor.row, 1) + ":" + convertToA1(blockEndCursor.row, cursor.col + T + 1));
        const backpropDimForLayout = String(allInputs[0]).includes("tinyDigits") ? maxDimTrimmed : maxDim;
        const blockrangeBack = abhSheet.getRange(convertToA1(blockStartCursor.row, cursor.col + T + 3) + ":" + convertToA1(blockEndCursor.row, cursor.col + T + 2 + T + (backpropDimForLayout > T ? backpropDimForLayout : T) + backpropDimForLayout + 6));
        if (backpropFlag && i == 0) {
            lrrow = blockStartCursor.row - 2;
            lrcol = cursor.col + 2 * T + 4 + (backpropDimForLayout > T ? backpropDimForLayout : T) + backpropDimForLayout + 6;
        }
        let fmt = blockrange.getFormat();
        let blockcolor = "#e0f0f0";
        if (i == 0) {
            verticalHeaderRange.getFormat().getFill().setColor("#305858");
        }
        else {
            if (i % 2 !== 0) {
                blockcolor = "#f1f4ed";
                verticalHeaderRange.getFormat().getFill().setColor("#6F7368");
            }
            else {
                verticalHeaderRange.getFormat().getFill().setColor("#305858");
            }
        }
        fmt.getFill().setColor(blockcolor);
        let borders = [
            ExcelScript.BorderIndex.edgeTop,
            ExcelScript.BorderIndex.edgeBottom,
            ExcelScript.BorderIndex.edgeLeft,
            ExcelScript.BorderIndex.edgeRight,
        ];
        abhSheet.getCell(cursor.row, cursor.col).setValue("");
        if (backpropFlag) {
            const fmt2 = blockrangeBack.getFormat();
            blockcolor = i < allOperations.length - 1 ? blockcolor : "white";
            if (i < allOperations.length - 1) {
                if (showCode) goBlockLinks(abhSheet, cursor.row, cursor.col + T + 3, codeLinkTarget, codeLinkLabel);
                else goTop(abhSheet, cursor.row, cursor.col + T + 3);
                backpropBlockSpans.push({
                    startRow: blockStartCursor.row,
                    endRow: blockEndCursor.row,
                    startCol: cursor.col + T + 3,
                    color: blockcolor
                });
            }
            fmt2.getFill().setColor(blockcolor);
        }
        cursor.row += 1;
        if (i < allOperations.length - 1) {
            makeArrow(abhSheet, "top2bottom", cursor.row, cursor.col, T);
        }
        dictVar["Input ⏵"] = `=HYPERLINK("#X", "🔗")`;
        if (hasLoss) {
            dictVar["OHE ⏵"] = `=HYPERLINK("#ohe", "🔗")`;
        }
        dictVar[`${i}| ${operation} ⏵`] = `=HYPERLINK("#${output_var_name}", "🔗"`;
        const _next_ops = values[i + 1][1] as string;
        if (_next_ops == "CELoss" || _next_ops == "") {
            endOfModel++;
        }
        if (endOfModel == 1) {
            allCode.push(")");
            allCode.push("");
            torchCode = genTorchCode("Predict", output_var_name.replace(/_\d+$/, ``), "X");
            allCode.push(torchCode);
            localCode.push(torchCode);
        }
    }
    const message = "📌 Blueprint Tab";
    linkedMessage(designSheet, message, `Blueprint!X`, 6, 6, true, 11, 2, "Black", "Purple");
    linkedMessage(abhSheet, `⏎`, `Architect!A1`, 0, 0, false, 15);
    let horHeadLeftRight_start = 1;
    let horHeadLeftRight_end = horHeadLeftRight_start + cursor.col + T;
    topHeads(abhSheet, "Forward ▼", horHeadLeftRight_start, horHeadLeftRight_end, "orange", 15, "black");
    let forEndCol = horHeadLeftRight_end;
    horHeadLeftRight_start = horHeadLeftRight_end + codeShiftRight - 2;
    horHeadLeftRight_end = horHeadLeftRight_start + 14;
    if (showCode) {
    }
    let oheStarts = abhSheet.getNamedItem("X");
    if (hasLoss) {
        oheStarts = abhSheet.getNamedItem("ohe");
    }
    const oheStarts_range = oheStarts.getRange();
    const rng = abhSheet.getNamedItem("X").getRange();
    const _keys = Object.keys(dictVar);
    const _values = Object.values(dictVar);
    let indStart = oheStarts_range.getRowIndex();
    let indStartRight = 0;
    indStartRight = cursor.col + T + 1;
    if (showList) {
        abhSheet.getRangeByIndexes(indStart, indStartRight + 3, 1, 1).getResizedRange(_keys.length - 1, 0).setFormulas(_keys.map(k => [k]));
        abhSheet.getRangeByIndexes(indStart, indStartRight + 4, 1, 1).getResizedRange(_values.length - 1, 0).setValues(_values.map(v => [v]));
        const rangeContents = abhSheet.getRangeByIndexes(indStart - 1, indStartRight, _keys.length + 1, 6);
        workbook.addNamedItem("Contents", rangeContents);
        rangeContents.getFormat().getFill().setColor("Moccasin");
        rangeContents.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
        rangeContents.getFormat().getFont().setSize(10);
        rangeContents.getFormat().getFont().setColor("DimGray");
        rangeContents.getFormat().getFont().setBold(false);
        const fmt = rangeContents.getFormat();
        const borders = [
            ExcelScript.BorderIndex.edgeTop,
            ExcelScript.BorderIndex.edgeBottom,
            ExcelScript.BorderIndex.edgeLeft,
            ExcelScript.BorderIndex.edgeRight,
        ];
        for (let b of borders) {
            fmt.getRangeBorder(b).setStyle(ExcelScript.BorderLineStyle.dashDotDot);
            fmt.getRangeBorder(b).setColor("orange");
        }
        abhSheet.getCell(indStart - 1, indStartRight).setValue("📝 List");
        abhSheet.getCell(indStart - 1, indStartRight)
            .getFormat().getFont().setBold(true);
        abhSheet.getCell(indStart - 1, indStartRight)
            .getFormat().getFont().setSize(12);
        abhSheet.getCell(indStart - 1, indStartRight)
            .getFormat().getFont().setUnderline(ExcelScript.RangeUnderlineStyle.none);
        abhSheet.getCell(indStart - 1, indStartRight)
            .getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    }
    cursor.row += 1;
    let lossName = "";
    let backpropStartCol = -1;
    let backpropStartRow = 0;
    if (backpropFlag) {
        const fKeys = Object.keys(forwardPass);
        let _formula = ``;
        let currentGradient = "";
        let prevGradient = "";
        let _eq = ``;
        const _updown = 0;
        let _leftright = cursor.col + T + indStartRight;
        _leftright = showList ? _leftright : cursor.col + T + (maxDim > T ? maxDim : T) + 6;
        if (String(allInputs[0]).includes("tinyDigits")) {
            _leftright = showList ? _leftright : cursor.col + T + (maxDimTrimmed > T ? maxDimTrimmed : T) + 6;
        }
        let maxDwRightCol = _leftright;
        for (let k = fKeys.length - 1; k > 0; k--) {
            const prevOpName = forwardPass[k - 1][1];
            if (prevOpName === "Linear") {
                const prevCols = forwardPass[k - 1][7];
                const weightCols = weightInformation[`W_${k - 1}`][4];
                maxDwRightCol = Math.max(maxDwRightCol, _leftright + prevCols + weightCols);
            }
        }
        lrcol = maxDwRightCol + 3;
        for (const block of backpropBlockSpans) {
            const fillWidth = lrcol - 3 - block.startCol + 2;
            if (fillWidth > 0) {
                abhSheet
                    .getRangeByIndexes(block.startRow, block.startCol, block.endRow - block.startRow + 1, fillWidth)
                    .getFormat()
                    .getFill()
                    .setColor(block.color);
            }
        }
        backpropStartCol = horHeadLeftRight_start;
        backpropStartRow = oheStarts_range.getRowIndex();
        headerFormat(abhSheet, "Learning rate, 𝜂 =", lrrow, lrcol, "right");
        const lrrange = abhSheet.getRangeByIndexes(lrrow, lrcol + 1, 1, 3);
        lrrange.merge();
        lrrange.getFormat().getFont().setName("Ink Free");
        lrrange.getFormat().getFont().setColor("white");
        lrrange.getFormat().getFont().setBold(true);
        lrrange.getFormat().getFill().setColor("CornflowerBlue");
        lrrange.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
        lrrange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
        lrrange.setValue(0.1);
        abhSheet.addNamedItem("lrate", lrrange);
        const regex_2 = /^(\w+),\s*dim\s*=\s*(\d+)$/i;
        const nums: number[] = [];
        nums.push(imRow * imCol);
        for (let i = 0; i < allOutputs.length; i++) {
            const cellValue = String(allOutputs[i]).trim();
            const match = cellValue.match(regex_2);
            if (match) {
                const num = Number(match[2]);
                nums.push(num);
            }
        }
        maxDim = Math.max(...nums);
        cursor.col = maxDim + 5;
        const lastDim = nums[nums.length - 1];
        topHeads(abhSheet, "Backprop ▲", horHeadLeftRight_start, lrcol - 2, "lightskyblue", 15, "black");
        for (let k = fKeys.length - 1; k > 0; k--) {
            let _inVarName = forwardPass[k][0];
            let _opsName = forwardPass[k][1];
            let _nextoperation = forwardPass[k][2];
            let _outVarName = forwardPass[k][3];
            let _posRow = forwardPass[k - 1][4];
            let _posCol = forwardPass[k - 1][5];
            let _nrows = forwardPass[k - 1][6];
            let _ncols = forwardPass[k - 1][7];
            let _nposRow = forwardPass[k][4];
            let _nposCol = forwardPass[k][5];
            let _nnrows = forwardPass[k][6];
            let _nncols = forwardPass[k][7];
            let _previnVarName = forwardPass[k - 1][0];
            let _prevopsName = forwardPass[k - 1][1];
            makeArrow(abhSheet, "bottom2top", _nposRow - 1, _leftright, T);
            if (k > 1) {
                makeArrow(abhSheet, "bottom2top", _posRow - 1, _leftright, T);
            }
            _posRow += 2;
            _posCol = 0;
            _nposRow += 2;
            _nposCol = 0;
            {
                prevGradient = toSubscript("G_" + String(k));
                if (_nextoperation == "MSELoss") {
                    lossName = _nextoperation;
                    currentGradient = "G_" + String(k);
                    _formula = `=(2/${lastDim})*(${_outVarName} - ohe)`;
                    _eq = `✏️${toSubscript("G_" + String(k))} = (2/𝘕)(${toSubscript(_outVarName)}−target)`;
                    headerFormat(abhSheet, `${toSubscript(currentGradient)}=Gradient(${toSubscript(_outVarName)})`, _nposRow - 2, _leftright + _nposCol);
                    writeEquation(abhSheet, _eq, _nposRow - 1, _leftright + _nposCol, "left");
                    const cell = abhSheet.getCell(_nposRow, _leftright + _nposCol);
                    cell.setFormula(_formula);
                    const range = doConditionalFormatting(abhSheet, _nposRow, _leftright + _nposCol, _nnrows, _nncols, "CornflowerBlue", "Tomato");
                    const cursor = { row: _nposRow - 1,
                        col: _leftright + _nposCol,
                        range: range };
                    updateName(workbook, cursor, currentGradient, abhSheet);
                }
                else if (_nextoperation == "CELoss") {
                    lossName = _nextoperation;
                    currentGradient = "G_" + String(k);
                    _formula = `=-ohe/${_outVarName}`;
                    _eq = `✏️${toSubscript("G_" + String(k))} = −${toSubscript(_outVarName)}/target`;
                    headerFormat(abhSheet, `${toSubscript(currentGradient)}=Gradient(${toSubscript(_outVarName)})`, _nposRow - 2, _leftright + _nposCol);
                    writeEquation(abhSheet, _eq, _nposRow - 1, _leftright + _nposCol, "left");
                    const cell = abhSheet.getCell(_nposRow, _leftright + _nposCol);
                    cell.setFormula(_formula);
                    const range = doConditionalFormatting(abhSheet, _nposRow, _leftright + _nposCol, _nnrows, _nncols, "CornflowerBlue", "Tomato");
                    const cursor = {
                        row: _nposRow - 1,
                        col: _leftright + _nposCol,
                        range: range
                    };
                    updateName(workbook, cursor, currentGradient, abhSheet);
                }
                prevGradient = toSubscript("G_" + String(k));
                if (_opsName == "Linear") {
                    const _wb = "W_" + String(k);
                    const _wbname = "(W_" + String(k) + "|" + "b_" + String(k) + ")";
                    _eq = `✏️ ${toSubscript(_wbname)}` + "ᵀ";
                    let nrows = weightInformation[_wb][3];
                    writeEquation(abhSheet, _eq, _posRow - 1, _leftright + _posCol - forwardPass[k][6] + nrows - 2, "right");
                    const cell = abhSheet.getCell(_posRow, _leftright + _posCol - forwardPass[k][6] - 1);
                    cell.setFormula(`=TRANSPOSE(${_wb})`);
                    const range = doConditionalFormatting(abhSheet, _posRow, _leftright + _posCol - forwardPass[k][6] - 1, _nrows + 1, forwardPass[k][6], "orange", "yellow", false, true);
                    const cursor = { row: _posRow - 1,
                        col: _leftright + _posCol - forwardPass[k][6] - 1,
                        range: range };
                    updateName(workbook, cursor, _wb + "T", abhSheet);
                    _formula = `=DROP(MMULT(${_wb}T, ${currentGradient}), -1)`;
                    const cell2 = abhSheet.getCell(_posRow, _leftright + _posCol);
                    cell2.setFormula(_formula);
                    makeArrow(abhSheet, "left2right", _posRow, _leftright + _posCol - 1, _nrows);
                    currentGradient = "G_" + String(k - 1);
                    headerFormat(abhSheet, `${toSubscript(currentGradient)}=Gradient(${prevGradient})`, _posRow - 2, _leftright + _posCol);
                    _eq = `✏️ ${toSubscript(currentGradient)}= ${toSubscript(_wb) + "ᵀ"} ${toSubscript("G_" + String(k))}`;
                    writeEquation(abhSheet, _eq, _posRow - 1, _leftright + _posCol, "left");
                    const range2 = doConditionalFormatting(abhSheet, _posRow, _leftright + _posCol, _nrows, _ncols, "CornflowerBlue", "Tomato");
                    const cursor2 = {
                        row: _posRow,
                        col: _leftright + _posCol,
                        range: range2
                    };
                    updateName(workbook, cursor2, currentGradient, abhSheet);
                }
                if (_opsName == "VA") {
                    _formula = `=MMULT(${currentGradient}, TRANSPOSE(${_inVarName}))`;
                    const cell = abhSheet.getCell(_posRow, _leftright + _posCol);
                    cell.setFormula(_formula);
                    const range2 = doConditionalFormatting(abhSheet, _posRow, _leftright + _posCol, _nrows, _ncols, "CornflowerBlue", "Tomato");
                    const cursor2 = {
                        row: _posRow,
                        col: _leftright + _posCol,
                        range: range2
                    };
                    updateName(workbook, cursor2, currentGradient, abhSheet);
                }
                else if (_opsName === "Tanh") {
                    let maskFormula = (_nextoperation === "Linear")
                        ? `=DROP(1-(${_outVarName}*${_outVarName}),-1)`
                        : `=1-(${_outVarName}*${_outVarName})`;
                    const cell = abhSheet.getCell(_posRow, _leftright + _posCol - _ncols - 1);
                    cell.setFormula(maskFormula);
                    const _tanhMask = "MaskTanh_" + String(k);
                    const _tanhMaskHeader = "𝒹_" + String(k);
                    _eq = `✏️ ${toSubscript(_tanhMaskHeader)}=1-${toSubscript(_outVarName)}²`;
                    writeEquation(abhSheet, _eq, _posRow - 1, _leftright + _posCol - _ncols - 1 + T - 1, "right");
                    const range = doConditionalFormatting(abhSheet, _posRow, _leftright + _posCol - _ncols - 1, _nrows, _ncols, "lightgray", "WhiteSmoke");
                    const cursor = {
                        row: _posRow,
                        col: _leftright + _posCol - _ncols - 1,
                        range: range
                    };
                    updateName(workbook, cursor, _tanhMask, abhSheet);
                    const cell2 = abhSheet.getCell(_posRow, _leftright + _posCol);
                    cell2.setFormula(`=${currentGradient}*${_tanhMask}`);
                    currentGradient = "G_" + String(k - 1);
                    makeArrow(abhSheet, "left2right", _posRow, _leftright + _posCol - 1, _nrows);
                    headerFormat(abhSheet, `${toSubscript(currentGradient)}=Gradient(${prevGradient})`, _posRow - 2, _leftright + _posCol);
                    _eq = `✏️ ${toSubscript(currentGradient)}=${toSubscript(_tanhMaskHeader)}⊙${toSubscript(_tanhMaskHeader)}`;
                    writeEquation(abhSheet, _eq, _posRow - 1, _leftright + _posCol, "left");
                    const range2 = doConditionalFormatting(abhSheet, _posRow, _leftright + _posCol, _nrows, _ncols, "CornflowerBlue", "Tomato");
                    const cursor2 = {
                        row: _posRow,
                        col: _leftright + _posCol,
                        range: range2
                    };
                    updateName(workbook, cursor2, currentGradient, abhSheet);
                }
                else if (_opsName === "GELU") {
                    let _outVarName_part = ``;
                    if (_nextoperation === "Linear")
                        _outVarName_part = `DROP(${_outVarName},-1)`;
                    else
                        _outVarName_part = _outVarName;
                    _formula = `=0.5*erf_${k} + 0.399*${_outVarName_part}* EXP(-0.5*${_outVarName_part}^2)`;
                    const cell = abhSheet.getCell(_posRow, _leftright + _posCol - _ncols - 1);
                    cell.setFormula(_formula);
                    const _geluMask = "MaskGELU_" + String(k);
                    const _geluMaskHeader = "𝒹_" + String(k);
                    _eq = `✏️ ${toSubscript(_geluMaskHeader)}=${toSubscript("u_" + String(k))}/2+${toSubscript(_outVarName)}⊙𝑒(-${toSubscript(_outVarName)}²/2)/√(2𝜋)`;
                    writeEquation(abhSheet, _eq, _posRow - 1, _leftright + _posCol - _ncols - 1 + T - 1, "right");
                    workbook.addNamedItem(_geluMask, cell);
                    const range = doConditionalFormatting(abhSheet, _posRow, _leftright + _posCol - _ncols - 1, _nrows, _ncols, "lightgray", "WhiteSmoke");
                    const cursor = {
                        row: _posRow,
                        col: _leftright + _posCol - _ncols - 1,
                        range: range
                    };
                    updateName(workbook, cursor, _geluMask, abhSheet);
                    if (_nextoperation === "Linear")
                        _formula = `=IF(DROP(${_outVarName},1)>0,${currentGradient},0)`;
                    else
                        _formula = `=IF(${_outVarName}>0,${currentGradient},0)`;
                    _formula = `=${currentGradient}* ${_geluMask}`;
                    const cell2 = abhSheet.getCell(_posRow, _leftright + _posCol);
                    cell2.setFormula(_formula);
                    makeArrow(abhSheet, "left2right", _posRow, _leftright + _posCol - 1, _nrows);
                    currentGradient = "G_" + String(k - 1);
                    headerFormat(abhSheet, `${toSubscript(currentGradient)}=Gradient(${prevGradient})`, _posRow - 2, _leftright + _posCol);
                    _eq = `✏️  ${toSubscript(currentGradient)} = ${toSubscript("G_" + String(k))}⊙${toSubscript(_geluMaskHeader)}`;
                    writeEquation(abhSheet, _eq, _posRow - 1, _leftright + _posCol, "left");
                    const range2 = doConditionalFormatting(abhSheet, _posRow, _leftright + _posCol, _nrows, _ncols, "CornflowerBlue", "Tomato");
                    const cursor2 = {
                        row: _posRow,
                        col: _leftright + _posCol,
                        range: range2
                    };
                    updateName(workbook, cursor2, currentGradient, abhSheet);
                }
                else if (_opsName == "LayerNorm") {
                    const stdName = `std_${_inVarName}`;
                    const leftCol = _leftright + _posCol - _ncols - 1;
                    const baseRow = _posRow;
                    const meanGName = `mG_${k}`;
                    const cellMeanG = abhSheet.getCell(baseRow, leftCol);
                    cellMeanG.setFormula(`=BYCOL(${currentGradient},LAMBDA(c,AVERAGE(c)))`);
                    _eq = `✏️ 𝓂 = mean(${toSubscript(`G_${k}`)})`;
                    writeEquation(abhSheet, _eq, _posRow - 1, leftCol, "left");
                    workbook.addNamedItem(meanGName, cellMeanG);
                    updateName(workbook, { row: baseRow, col: leftCol, range: cellMeanG.getResizedRange(0, _ncols - 1) }, meanGName, abhSheet);
                    doConditionalFormatting(abhSheet, baseRow, leftCol, 1, _ncols, "lightgray", "WhiteSmoke");
                    const meanGYName = `mGY_${k}`;
                    const cellMeanGY = abhSheet.getCell(baseRow + 2, leftCol);
                    if (_nextoperation === "Linear")
                        _formula = `=BYCOL(${currentGradient}*DROP(${_outVarName},-1), LAMBDA(c,AVERAGE(c)))`;
                    else
                        _formula = `=BYCOL(${currentGradient}*${_outVarName}, LAMBDA(c,AVERAGE(c)))`;
                    cellMeanGY.setFormula(_formula);
                    _eq = `✏️ 𝘬 = mean(${toSubscript(`G_${k}`)} ⊙ ${toSubscript(`${_outVarName}`)})`;
                    writeEquation(abhSheet, _eq, baseRow + 1, leftCol, "left");
                    workbook.addNamedItem(meanGYName, cellMeanGY);
                    updateName(workbook, { row: baseRow + 2, col: leftCol, range: cellMeanGY.getResizedRange(0, _ncols - 1) }, meanGYName, abhSheet);
                    doConditionalFormatting(abhSheet, baseRow + 2, leftCol, 1, _ncols, "lightgray", "WhiteSmoke");
                    const dXName = `G_${k - 1}`;
                    const cellDX = abhSheet.getCell(_posRow, _leftright + _posCol);
                    if (_nextoperation === "Linear")
                        _formula = `=DROP((${currentGradient} - ${meanGName} - ${_outVarName}*${meanGYName}) * 1/${stdName}, -1)`;
                    else
                        _formula = `=(${currentGradient} - ${meanGName} - ${_outVarName}*${meanGYName}) * 1/${stdName}`;
                    cellDX.setFormula(_formula);
                    makeArrow(abhSheet, "left2right", _posRow, _leftright + _posCol - 1, 1);
                    makeArrow(abhSheet, "left2right", _posRow + 2, _leftright + _posCol - 1, 1);
                    headerFormat(abhSheet, `${toSubscript(dXName)}=Gradient(${prevGradient})`, _posRow - 2, _leftright + _posCol);
                    currentGradient = dXName;
                    _eq = `✏️ ${toSubscript(currentGradient)} = (${toSubscript("G_" + String(k))}−𝓂−𝘬 ⊙ ${toSubscript(_outVarName)})/${toSubscript("𝜎²_" + String(k))}`;
                    writeEquation(abhSheet, _eq, _posRow - 1, _leftright + _posCol, "left");
                    const rangeDX = doConditionalFormatting(abhSheet, _posRow, _leftright + _posCol, _nrows, _ncols, "CornflowerBlue", "Tomato");
                    updateName(workbook, { row: _posRow, col: _leftright + _posCol, range: rangeDX }, dXName, abhSheet);
                }
                else if (_opsName === "Softmax" && _prevopsName != "CELoss") {
                    const leftCol = _leftright + _posCol - _ncols - 1;
                    let baseRow = _posRow;
                    let formula_ = ``;
                    formula_ =
                        (_nextoperation === "Linear")
                            ? `=BYCOL(${currentGradient}*DROP(${_outVarName},-1) ,LAMBDA(c,SUM(c)))`
                            : `=BYCOL(${currentGradient}*${_outVarName} ,LAMBDA(c,SUM(c)))`;
                    const colSumName = `sumGX_${k}`;
                    const cellColSum = abhSheet.getCell(baseRow, leftCol);
                    cellColSum.setFormula(formula_);
                    _eq = `✏️  ${toSubscript("s_" + String(k))} = ` + toSubscript(`Σ(G_${k}`) + "⊙" + toSubscript(`${_outVarName})`);
                    writeEquation(abhSheet, _eq, baseRow - 1, leftCol + T - 1, "right");
                    workbook.addNamedItem(colSumName, cellColSum);
                    updateName(workbook, { row: baseRow + _nrows + 1, col: leftCol, range: cellColSum.getResizedRange(0, _ncols - 1) }, colSumName, abhSheet);
                    doConditionalFormatting(abhSheet, baseRow, leftCol, 1, _ncols, "lightgray", "WhiteSmoke");
                    formula_ =
                        (_nextoperation === "Linear")
                            ? `=DROP(${_outVarName}*(${currentGradient}-${colSumName}), -1)` :
                            `=${_outVarName}*(${currentGradient}-${colSumName})`;
                    const cellGZ = abhSheet.getCell(_posRow, _leftright + _posCol);
                    cellGZ.setFormula(formula_);
                    makeArrow(abhSheet, "left2right", _posRow, _leftright + _posCol - 1, 1);
                    currentGradient = `G_${k - 1}`;
                    headerFormat(abhSheet, `${toSubscript(currentGradient)}=Gradient(${prevGradient})`, _posRow - 2, _leftright + _posCol);
                    _eq = `✏️  ${toSubscript(currentGradient)} = ${toSubscript(_outVarName)}⊙(${prevGradient}−${toSubscript("s_" + String(k))})`;
                    writeEquation(abhSheet, _eq, _posRow - 1, _leftright + _posCol, "left");
                    const gzRange = doConditionalFormatting(abhSheet, _posRow, _leftright + _posCol, _nrows, _ncols, "CornflowerBlue", "Tomato");
                    updateName(workbook, { row: _posRow, col: _leftright + _posCol, range: gzRange }, currentGradient, abhSheet);
                }
                else if (_opsName == "ReLU") {
                    if (_nextoperation === "Linear")
                        _formula = `=DROP(IF(${_outVarName}>0, 1,0),-1)`;
                    else
                        _formula = `=IF(${_outVarName}>0, 1,0)`;
                    const cell = abhSheet.getCell(_posRow, _leftright + _posCol - _ncols - 1);
                    cell.setFormula(_formula);
                    const _reluMask = "MaskReLU_" + String(k);
                    const _reluMaskHeader = "𝒹_" + String(k);
                    _eq = `✏️  ${toSubscript(_reluMaskHeader)} = 1 if ${toSubscript(_outVarName)}>0; else 0`;
                    writeEquation(abhSheet, _eq, _posRow - 1, _leftright + _posCol - _ncols - 1 + T - 1, "right");
                    workbook.addNamedItem(_reluMask, cell);
                    const range = doConditionalFormatting(abhSheet, _posRow, _leftright + _posCol - _ncols - 1, _nrows, _ncols, "lightgray", "WhiteSmoke");
                    const cursor = {
                        row: _posRow,
                        col: _leftright + _posCol - _ncols - 1,
                        range: range
                    };
                    updateName(workbook, cursor, _reluMask, abhSheet);
                    if (_nextoperation === "Linear")
                        _formula = `=IF(DROP(${_outVarName},1)>0,${currentGradient},0)`;
                    else
                        _formula = `=IF(${_outVarName}>0,${currentGradient},0)`;
                    _formula = `=${currentGradient} *${_reluMask}`;
                    const cell2 = abhSheet.getCell(_posRow, _leftright + _posCol);
                    cell2.setFormula(_formula);
                    makeArrow(abhSheet, "left2right", _posRow, _leftright + _posCol - 1, _nrows);
                    currentGradient = "G_" + String(k - 1);
                    headerFormat(abhSheet, `${toSubscript(currentGradient)}=Gradient(${prevGradient})`, _posRow - 2, _leftright + _posCol);
                    _eq = `✏️  ${toSubscript(currentGradient)} = ${toSubscript("G_" + String(k))}⊙${toSubscript(_reluMaskHeader)}`;
                    writeEquation(abhSheet, _eq, _posRow - 1, _leftright + _posCol, "left");
                    const range2 = doConditionalFormatting(abhSheet, _posRow, _leftright + _posCol, _nrows, _ncols, "CornflowerBlue", "Tomato");
                    const cursor2 = {
                        row: _posRow,
                        col: _leftright + _posCol,
                        range: range2
                    };
                    updateName(workbook, cursor2, currentGradient, abhSheet);
                }
                else if (_opsName == "Attention") {
                    const _v = "V_" + String(k);
                    const cell = abhSheet.getCell(_posRow, _leftright + _posCol - forwardPass[k][6] - 1);
                    cell.setFormula(`=TRANSPOSE(${_v})`);
                    const range = doConditionalFormatting(abhSheet, _posRow, _leftright + _posCol - forwardPass[k][6] - 1, _nrows + 1, forwardPass[k][6], "orange", "yellow", false, true);
                    const cursor = {
                        row: _posRow - 1,
                        col: _leftright + _posCol - forwardPass[k][6] - 1,
                        range: range
                    };
                    updateName(workbook, cursor, `${_v}T`);
                    _formula = `=DROP(MMULT(${_v}T, ${currentGradient}), -1)`;
                    const cell2 = abhSheet.getCell(_posRow, _leftright + _posCol);
                    cell2.setFormula(_formula);
                    currentGradient = "G_" + String(k - 1);
                    headerFormat(abhSheet, toSubscript(currentGradient), _posRow - 1, _leftright + _posCol);
                    const range2 = doConditionalFormatting(abhSheet, _posRow, _leftright + _posCol, _nrows, _ncols, "CornflowerBlue", "Tomato");
                    const cursor2 = {
                        row: _posRow,
                        col: _leftright + _posCol,
                        range: range2
                    };
                    updateName(workbook, cursor2, currentGradient);
                }
            }
            if (_prevopsName == "Linear") {
                _formula = `=MMULT(${currentGradient}, TRANSPOSE(${_previnVarName}))`;
                const cell2 = abhSheet.getCell(_posRow, _leftright + _posCol + _ncols + 1);
                cell2.setFormula(_formula);
                makeArrow(abhSheet, "left2right", _posRow, _leftright + _posCol + _ncols, _nrows);
                let ncols = weightInformation[`W_${k - 1}`][4];
                headerFormat(abhSheet, toSubscript("𝒹W_" + String(k - 1)), _posRow - 2, _leftright + _posCol + _ncols + ncols);
                _eq = `✏️  ${toSubscript("𝒹W_" + String(k - 1))} = ${toSubscript(currentGradient)}⊙${toSubscript(_previnVarName)}ᵀ`;
                writeEquation(abhSheet, _eq, _posRow - 1, _leftright + _posCol + _ncols + ncols, "right");
                const rng = abhSheet.getNamedItem(_previnVarName).getRange();
                const range2 = doConditionalFormatting(abhSheet, _posRow, _leftright + _posCol + _ncols + 1, _nrows, ncols, "orange", "yellow", false, true);
                const cursor2 = {
                    row: _posRow,
                    col: _leftright + _posCol + _ncols + 1,
                    range: range2
                };
                updateName(workbook, cursor2, "dW_" + String(k - 1), abhSheet);
                _formula = `=W_${k - 1} - lrate * dW_${k - 1} `;
                const header = toSubscript("W'_" + String(k - 1));
                headerFormat(abhSheet, header, _posRow - 2, lrcol);
                _eq = `✏️ ${toSubscript("W'_" + String(k - 1))} = ${toSubscript("W_" + String(k - 1))} − 𝜂.${toSubscript("𝒹W_" + String(k - 1))}`;
                writeEquation(abhSheet, _eq, _posRow - 1, lrcol, "left");
                const cell3 = abhSheet.getCell(_posRow, lrcol);
                cell3.setFormula(_formula);
                makeArrow(abhSheet, "left2right", _posRow, lrcol - 1, _nrows);
                const range3 = doConditionalFormatting(abhSheet, _posRow, lrcol, _nrows, ncols, "orange", "yellow", false, true);
                const cursor3 = {
                    row: _posRow,
                    col: lrcol,
                    range: range3
                };
                weightUpdateRightCol = Math.max(weightUpdateRightCol, lrcol + ncols - 1);
                updateName(workbook, cursor3, "newW_" + String(k - 1), abhSheet);
            }
            doneMessage(k + 4, designSheet, 3, true);
        }
        horHeadLeftRight_start = lrcol;
        const weightUpdateHeaderEnd = Math.max(weightUpdateRightCol, horHeadLeftRight_start + maxDim);
        topHeads(abhSheet, "Weight Update", horHeadLeftRight_start, weightUpdateHeaderEnd, "darkseagreen", 15, "black");
    }
    if (showCode && seqLayers.length > 0) {
        const backpropCodeAnchor = Math.max(weightUpdateRightCol, lrcol + maxDim);
        const codeStartCol = backpropFlag
            ? backpropCodeAnchor + codeShiftRight
            : cursor.col + T + codeShiftRight;
        const codeEndCol = codeStartCol + 25;
        generatedCodeStartCol = codeStartCol;
        generatedCodeColCount = codeEndCol - generatedCodeStartCol + 1;
        topHeads(abhSheet, "PyTorch Code", codeStartCol - 1, codeEndCol, "plum", 15, "black");
        const fullCode: string[] = [];
        const fullCodeLinks: string[] = [];
        const fullCodeAnchors: string[][] = [];
        function pushFullCodeLine(line: string, link: string = "", anchors: string[] = []): void {
            fullCode.push(line);
            fullCodeLinks.push(link);
            fullCodeAnchors.push(anchors);
        }
        pushFullCodeLine("import torch");
        pushFullCodeLine("import torch.nn as nn");
        pushFullCodeLine("");
        pushFullCodeLine(inputTensorCode || `X = torch.randn(${inputBatch}, ${inputFeatures})`, "X");
        if (lossType === "CELoss") {
            pushFullCodeLine(`target = torch.randint(0, ${nClasses}, (${inputBatch},))`, "ohe");
        }
        else if (lossType === "MSELoss") {
            const outDim = lastLinearOut > 0 ? lastLinearOut : inputFeatures;
            pushFullCodeLine(`target = torch.randn(${inputBatch}, ${outDim})`, "ohe");
        }
        pushFullCodeLine("");
        pushFullCodeLine("model = nn.Sequential(");
        for (let _si = 0; _si < seqLayers.length; _si++) {
            const comma = _si < seqLayers.length - 1 ? "," : "";
            pushFullCodeLine(seqLayers[_si] + comma, seqLayerLinks[_si] || "", seqLayerAnchors[_si] || []);
        }
        pushFullCodeLine(")");
        pushFullCodeLine("");
        if (lossType === "CELoss") {
            pushFullCodeLine("loss_fn = nn.CrossEntropyLoss()", lossVarName, lossFnAnchorNames);
        }
        else if (lossType === "MSELoss") {
            pushFullCodeLine("loss_fn = nn.MSELoss()", lossVarName, lossFnAnchorNames);
        }
        pushFullCodeLine("# Forward pass");
        pushFullCodeLine("X_out = model(X)", modelOutputLinkName || (seqLayerLinks.length > 0 ? (seqLayerLinks[seqLayerLinks.length - 1] || "") : ""));
        if (lossType !== "") {
            pushFullCodeLine("L = loss_fn(X_out, target)", lossVarName, lossEvalAnchorNames);
        }
        if (backpropFlag && hasLoss && lossType !== "") {
            const generatedEpochCount = Number(numOfEpoch) + 1;
            pushFullCodeLine("");
            pushFullCodeLine("# === Training Loop ===");
            pushFullCodeLine("num_epochs = " + String(generatedEpochCount));
            pushFullCodeLine("optimizer = torch.optim.SGD(");
            pushFullCodeLine("    model.parameters(), lr=0.1)");
            pushFullCodeLine("");
            pushFullCodeLine("for epoch in range(num_epochs):");
            pushFullCodeLine("    X_out = model(X)", modelOutputLinkName || (seqLayerLinks.length > 0 ? (seqLayerLinks[seqLayerLinks.length - 1] || "") : ""));
            pushFullCodeLine("    L = loss_fn(X_out, target)", lossVarName);
            pushFullCodeLine("    optimizer.zero_grad()  # reset dW");
            pushFullCodeLine("    L.backward()           # dL/dW");
            pushFullCodeLine("    optimizer.step()       # W -= lr * dW");
        }
        const writeRow = codeTopRow > 0 ? codeTopRow : 1;
        writeTorchCode(abhSheet, "", fullCode, writeRow, codeStartCol, undefined, undefined, undefined, undefined, fullCodeLinks, fullCodeAnchors);
        const torchCodeRange = abhSheet.getRangeByIndexes(writeRow, generatedCodeStartCol, fullCode.length, generatedCodeColCount);
        updateName(workbook, { row: writeRow, col: generatedCodeStartCol, range: torchCodeRange }, "torchcode", abhSheet);
    }
    if (showCode) {
        const message2 = "📌 PyTorch Code";
        linkedMessage(designSheet, message2, `Blueprint!torchcode`, 7, 6, true, 11, 2, "Blue", "Purple");
    }
    if (backpropFlag) {
        const weightKeys = Object.keys(weightInformation);
        let weightFormula = ``;
        let lrFormula = ``;
        let sourceName = "Blueprint";
        for (let ep = 1; ep <= Number(numOfEpoch); ep++) {
            const copiedName = `Blueprint_ep${ep}`;
            const sourceSheet = workbook.getWorksheet(sourceName);
            let copiedSheet = workbook.getWorksheet(copiedName);
            copiedSheet = sourceSheet.copy(ExcelScript.WorksheetPositionType.after, sourceSheet);
            copiedSheet.setTabColor("Green");
            copiedSheet.setName(copiedName);
            desighedText(copiedSheet, `Epoch-${ep}`, 0, 10, 1, 1, "white", "Ink Free", 13, "Tomato", true, "left", "center");
            if (generatedCodeStartCol >= 0 && generatedCodeColCount > 0) {
                const deleteCodeStartCol = Math.max(0, generatedCodeStartCol - 1);
                const deleteCodeColCount = generatedCodeColCount + (generatedCodeStartCol > 0 ? 1 : 0);
                copiedSheet
                    .getRangeByIndexes(0, deleteCodeStartCol, 1, deleteCodeColCount)
                    .getEntireColumn()
                    .delete(ExcelScript.DeleteShiftDirection.left);
            }
            for (const w of weightKeys) {
                let layerno = weightInformation[w][0];
                let wrow = weightInformation[w][1];
                let wcol = weightInformation[w][2];
                let src_var = `${sourceName}!new${w}`;
                let trgt_var = `${copiedName}!${w}`;
                weightFormula = `=${src_var}`;
                const targetNameObj = copiedSheet.getNamedItem(w);
                const targetRange = targetNameObj.getRange();
                targetRange.clear(ExcelScript.ClearApplyTo.contents);
                copiedSheet
                    .getCell(wrow, wcol)
                    .setFormula(weightFormula);
            }
            lrFormula = `=Blueprint!lrate `;
            copiedSheet
                .getCell(lrrow, lrcol + 1)
                .setFormula(lrFormula);
            let message = `Epoch ${ep}`;
            message = "📌 Blueprint " + message + " Tab";
            linkedMessage(designSheet, message, `${copiedName}!X`, 7 + ep, 6, true, 11, 2, "Green", "PaleGreen");
            linkedMessage(copiedSheet, `⏎`, `Architect!A1`, 0, 0, false, 15);
            const sourceNames = sourceSheet.getNames();
            sourceName = copiedName;
            designSheet.activate();
        }
        const finalTrainingSheet = workbook.getWorksheet(sourceName);
        const inferenceName = "Blueprint_inference";
        let inferenceSheet = workbook.getWorksheet(inferenceName);
        if (inferenceSheet) {
            inferenceSheet.delete();
        }
        inferenceSheet = finalTrainingSheet.copy(ExcelScript.WorksheetPositionType.after, finalTrainingSheet);
        inferenceSheet.setTabColor("Orange");
        inferenceSheet.setName(inferenceName);
        desighedText(inferenceSheet, "Inference", 0, 10, 1, 1, "white", "Ink Free", 13, "DarkOrange", true, "left", "center");
        if (generatedCodeStartCol >= 0 && generatedCodeColCount > 0) {
            const deleteCodeStartCol = Math.max(0, generatedCodeStartCol - 1);
            const deleteCodeColCount = generatedCodeColCount + (generatedCodeStartCol > 0 ? 1 : 0);
            inferenceSheet
                .getRangeByIndexes(0, deleteCodeStartCol, 1, deleteCodeColCount)
                .getEntireColumn()
                .delete(ExcelScript.DeleteShiftDirection.left);
        }
        if (backpropStartCol >= 0) {
            const clearColCount = horHeadLeftRight_end;
            if (clearColCount > 0) {
                const clearRange = inferenceSheet.getRangeByIndexes(indStart, backpropStartCol, 500, horHeadLeftRight_start + maxDim);
                clearRange.unmerge();
                clearRange.clear(ExcelScript.ClearApplyTo.contents);
                clearRange.clear(ExcelScript.ClearApplyTo.formats);
                const clearRangeUp = inferenceSheet.getRangeByIndexes(0, backpropStartCol, 2, horHeadLeftRight_start + maxDim);
                clearRangeUp.unmerge();
                clearRangeUp.clear(ExcelScript.ClearApplyTo.contents);
                clearRangeUp.clear(ExcelScript.ClearApplyTo.formats);
            }
        }
        linkedMessage(designSheet, "📌 Blueprint Inference Tab", `${inferenceName}!X`, 7 + Number(numOfEpoch) + 1, 6, true, 11, 2, "DarkOrange", "Moccasin");
        linkedMessage(inferenceSheet, `⏎`, `Architect!A1`, 0, 0, false, 15);
        designSheet.activate();
        dashboardSheet = workbook.addWorksheet("Dashboard");
        dashboardSheet.setPosition(Number(numOfEpoch) + 3);
        dashboardSheet.setTabColor("Purple");
        dashboardSheet.setShowGridlines(false);
        dashboardSheet.setShowHeadings(false);
        dashboardSheet.getRange().getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
        dashboardSheet.getRange().getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
        dashboardSheet.getRange().getFormat().setColumnWidth(W);
        dashboardSheet.getRange().getFormat().setRowHeight(H);
        dashboardSheet.getRange().getFormat().getFont().setName("Ink Free");
        dashboardSheet.getRange().getFormat().getFont().setSize(7);
        dashboardSheet.getRange().getFormat().getFont().setBold(false);
        titleMaker(dashboardSheet, "HackPropagation · Dashboard", 0, 1);
        const dashboardListName = "dashboard_list";
        const lossAnchorName = "dashboard_loss";
        const weightAnchorName = "dashboard_weights";
        const gradientAnchorName = "dashboard_gradients";
        const summaryAnchorName = "dashboard_grad_summary";
        const dashboardListRef = `Dashboard!${dashboardListName}`;
        let nrows = 0;
        let ncols = 0;
        const listBlockRow = 3;
        const listBlockCol = 1;
        const listBlockHeight = 8;
        const listBlockWidth = 8;
        addDashboardAnchor(workbook, dashboardSheet, dashboardListName, listBlockRow, listBlockCol);
        const lossBlockRow = 3;
        const lossBlockCol = 10;
        const lossBlockHeight = 17;
        const lossBlockWidth = 24;
        const lossStartrow = lossBlockRow + 2;
        const lossStartcol = lossBlockCol + 1;
        const seqFromula = `=SEQUENCE(${Number(numOfEpoch) + 1}, 1, 0)`;
        const formula1 = `='Blueprint'!LossMean`;
        desighedText(dashboardSheet, "Epochs", lossStartrow - 1, lossStartcol - 1, 1, 1);
        desighedText(dashboardSheet, "Loss", lossStartrow - 1, lossStartcol, 1, 1);
        dashboardSheet.getCell(lossStartrow, lossStartcol - 1).setFormula(seqFromula);
        dashboardSheet.getCell(lossStartrow, lossStartcol).setFormula(formula1);
        for (let ep = 1; ep <= Number(numOfEpoch); ep++) {
            const epochName = `Blueprint_ep${ep}`;
            const formula = `='` + epochName + `'!LossMean`;
            dashboardSheet.getRangeByIndexes(lossStartrow + ep, lossStartcol, 1, 1).setFormula(formula);
        }
        doConditionalFormatting(dashboardSheet, lossStartrow, lossStartcol, Number(numOfEpoch) + 1, 1, "orange", "yellow", false, true);
        genChart(dashboardSheet, "Loss", lossStartrow, lossStartcol, Number(numOfEpoch) + 1, 1, lossBlockCol + 3, lossBlockRow + 1, lossName, "Epoch");
        drawDashboardBlock(dashboardSheet, lossBlockRow, lossBlockCol, lossBlockHeight, lossBlockWidth, "Loss", dashboardListRef, "Black");
        addDashboardAnchor(workbook, dashboardSheet, lossAnchorName, lossBlockRow, lossBlockCol);
        const gradientBlockRow = 3;
        const gradientBlockCol = lossBlockCol + lossBlockWidth + 5;
        const gradientTopRow = gradientBlockRow + 2;
        const gradientTopCol = gradientBlockCol + 1;
        let gradientDataCol = gradientBlockCol + 2;
        let gradientMaxBottomRow = gradientTopRow + 3;
        desighedText(dashboardSheet, "Gradients", gradientTopRow, gradientTopCol + 1, 1, 1, "white", "Ink Free", 12, "Gray", true, "left", "bottom");
        desighedText(dashboardSheet, "⟶ ⟶ ⟶ ⟶", gradientTopRow + 1, gradientTopCol + 1, 1, 1, "white", "Ink Free", 12, "Black", false, "left", "bottom");
        desighedText(dashboardSheet, "Epochs", gradientTopRow + 2, gradientTopCol - 1, 1, 1, "white", "Ink Free", 12, "Gray", true, "right", "top");
        desighedText(dashboardSheet, "↓", gradientTopRow + 3, gradientTopCol - 1, 1, 1, "white", "Ink Free", 12, "Black", false, "right", "top");
        desighedText(dashboardSheet, "↓", gradientTopRow + 4, gradientTopCol - 1, 1, 1, "white", "Ink Free", 12, "Black", false, "right", "top");
        desighedText(dashboardSheet, "↓", gradientTopRow + 5, gradientTopCol - 1, 1, 1, "white", "Ink Free", 12, "Black", false, "right", "top");
        let _keys = Object.keys(layerInformation);
        let gradientEndCol = gradientDataCol;
        for (let layers = _keys.length - 2; layers > 0; layers--) {
            let rowgaps = gradientTopRow + 3;
            for (let ep = 0; ep <= Number(numOfEpoch); ep++) {
                headerFormat(dashboardSheet, `${toSubscript("G_" + String(layers))} (ep:${ep})`, rowgaps - 1, gradientDataCol);
                let wgt = ep == 0 ? `Blueprint!G_${layers - 1}` : `Blueprint_ep${ep}!G_${layers - 1}`;
                weightFormula = `=${wgt}`;
                dashboardSheet.getCell(rowgaps, gradientDataCol).setFormula(weightFormula);
                nrows = layerInformation[layers][1];
                doConditionalFormatting(dashboardSheet, rowgaps, gradientDataCol, nrows, T, "CornflowerBlue", "Tomato", true, false);
                rowgaps += nrows + 1;
                gradientMaxBottomRow = Math.max(gradientMaxBottomRow, rowgaps - 1);
            }
            gradientDataCol += T + 1;
            gradientEndCol = gradientDataCol - 1;
        }
        const gradientBlockHeight = Math.max(14, gradientMaxBottomRow - gradientBlockRow + 2);
        const gradientBlockWidth = Math.max(18, gradientEndCol - gradientBlockCol + 2);
        drawDashboardBlock(dashboardSheet, gradientBlockRow, gradientBlockCol, gradientBlockHeight, gradientBlockWidth, "Gradient Flow", dashboardListRef, "Black");
        addDashboardAnchor(workbook, dashboardSheet, gradientAnchorName, gradientBlockRow, gradientBlockCol);
        const dashboardBlockGap = 4;
        const secondRowTop = Math.max(lossBlockRow + lossBlockHeight, gradientBlockRow + gradientBlockHeight) + dashboardBlockGap;
        const weightBlockRow = secondRowTop;
        const weightBlockCol = lossBlockCol;
        const weightTopRow = weightBlockRow + 2;
        const weightTopCol = weightBlockCol + 1;
        let weightDataCol = weightTopCol + 1;
        let weightMaxBottomRow = weightTopRow + 3;
        desighedText(dashboardSheet, "Weights", weightTopRow, weightTopCol + 1, 1, 1, "white", "Ink Free", 12, "Gray", true, "left", "bottom");
        desighedText(dashboardSheet, "⟶ ⟶ ⟶ ⟶", weightTopRow + 1, weightTopCol + 1, 1, 1, "white", "Ink Free", 12, "Black", false, "left", "bottom");
        desighedText(dashboardSheet, "Epochs", weightTopRow + 2, weightTopCol - 1, 1, 1, "white", "Ink Free", 12, "Gray", true, "right", "top");
        desighedText(dashboardSheet, "↓", weightTopRow + 3, weightTopCol - 1, 1, 1, "white", "Ink Free", 12, "Black", false, "right", "top");
        desighedText(dashboardSheet, "↓", weightTopRow + 4, weightTopCol - 1, 1, 1, "white", "Ink Free", 12, "Black", false, "right", "top");
        desighedText(dashboardSheet, "↓", weightTopRow + 5, weightTopCol - 1, 1, 1, "white", "Ink Free", 12, "Black", false, "right", "top");
        for (const w of weightKeys) {
            let rowgaps = weightTopRow + 3;
            let layerno = weightInformation[w][0];
            nrows = weightInformation[w][3];
            ncols = weightInformation[w][4];
            for (let ep = 0; ep <= Number(numOfEpoch); ep++) {
                headerFormat(dashboardSheet, `${toSubscript("𝒹W_" + String(layerno))} (ep:${ep})`, rowgaps - 1, weightDataCol);
                let wgt = ep == 0 ? `Blueprint!d${w}` : `Blueprint_ep${ep}!d${w}`;
                weightFormula = `=${wgt}`;
                dashboardSheet.getCell(rowgaps, weightDataCol).setFormula(weightFormula);
                doConditionalFormatting(dashboardSheet, rowgaps, weightDataCol, nrows, ncols, "orange", "yellow", false, true);
                rowgaps += nrows + 1;
                weightMaxBottomRow = Math.max(weightMaxBottomRow, rowgaps - 1);
            }
            weightDataCol += ncols + 1;
        }
        const weightBlockHeight = Math.max(16, weightMaxBottomRow - weightBlockRow + 2);
        const weightBlockWidth = Math.max(22, weightDataCol - weightBlockCol + 1);
        drawDashboardBlock(dashboardSheet, weightBlockRow, weightBlockCol, weightBlockHeight, weightBlockWidth, "Weight Update", dashboardListRef, "Black");
        addDashboardAnchor(workbook, dashboardSheet, weightAnchorName, weightBlockRow, weightBlockCol);
        const summaryBlockRow = secondRowTop;
        const summaryBlockCol = Math.max(gradientBlockCol, weightBlockCol + weightBlockWidth + dashboardBlockGap);
        const gradientLayerCount = Math.max(0, Object.keys(layerInformation).length - 2);
        if (gradientLayerCount > 0) {
            const summaryStartRow = summaryBlockRow + 2;
            const summaryStartCol = summaryBlockCol + 1;
            const summaryDataStartRow = summaryStartRow + 1;
            const summaryDataStartCol = summaryStartCol;
            const summaryEpochCount = Math.max(1, Number(numOfEpoch) + 1);
            const summaryRowCount = summaryEpochCount + 1;
            const summaryColCount = gradientLayerCount + 1;
            headerFormat(dashboardSheet, "Ep", summaryStartRow, summaryStartCol);
            for (let gi = 0; gi < gradientLayerCount; gi++) {
                headerFormat(dashboardSheet, toSubscript("G_" + String(gi)), summaryStartRow, summaryStartCol + 1 + gi);
            }
            dashboardSheet.getCell(summaryDataStartRow, summaryDataStartCol).setFormula(`=SEQUENCE(${summaryEpochCount},1,0)`);
            for (let ep = 0; ep < summaryEpochCount; ep++) {
                for (let gi = 0; gi < gradientLayerCount; gi++) {
                    const gradRef = ep === 0 ? `Blueprint!G_${gi}` : `Blueprint_ep${ep}!G_${gi}`;
                    const formula = `=IFERROR(SUMPRODUCT(ABS(${gradRef}))/COUNT(${gradRef}),"")`;
                    dashboardSheet.getCell(summaryDataStartRow + ep, summaryDataStartCol + 1 + gi).setFormula(formula);
                }
            }
            doConditionalFormatting(dashboardSheet, summaryDataStartRow, summaryDataStartCol + 1, summaryEpochCount, gradientLayerCount, "#DCEBFF", "#FFB38A", true, true);
            const tableBottomRow = summaryDataStartRow + summaryEpochCount - 1;
            const summaryChartTop = tableBottomRow + 3;
            genMultiSeriesChart(dashboardSheet, "Mean gradient", summaryStartRow, summaryStartCol, summaryRowCount, summaryColCount, summaryBlockCol + 1, summaryChartTop, "Epoch", "Mean gradient");
            const summaryBlockHeight = summaryChartTop - summaryBlockRow + 18;
            const summaryBlockWidth = Math.max(18, summaryColCount + 10);
            drawDashboardBlock(dashboardSheet, summaryBlockRow, summaryBlockCol, summaryBlockHeight, summaryBlockWidth, "Mean Gradient", dashboardListRef, "Black");
            addDashboardAnchor(workbook, dashboardSheet, summaryAnchorName, summaryBlockRow, summaryBlockCol);
        }
        drawDashboardBlock(dashboardSheet, listBlockRow, listBlockCol, listBlockHeight, listBlockWidth, "Visualizations", "", "Black");
        linkedMessage(dashboardSheet, "Loss", `Dashboard!${lossAnchorName}`, listBlockRow + 1, listBlockCol + 1, true, 9, 4, "RoyalBlue");
        linkedMessage(dashboardSheet, "Gradient Flow", `Dashboard!${gradientAnchorName}`, listBlockRow + 2, listBlockCol + 1, true, 9, 4, "RoyalBlue");
        linkedMessage(dashboardSheet, "Weight Update", `Dashboard!${weightAnchorName}`, listBlockRow + 3, listBlockCol + 1, true, 9, 4, "RoyalBlue");
        linkedMessage(dashboardSheet, "Mean gradient", `Dashboard!${summaryAnchorName}`, listBlockRow + 4, listBlockCol + 1, true, 9, 4, "RoyalBlue");
        linkedMessage(designSheet, `📌 Dashboard Tab`, `Dashboard!A1`, 7 + Number(numOfEpoch) + 2, 6, true, 11, 2, "Purple", "OceanBlue");
        linkedMessage(dashboardSheet, `⏎`, `Architect!A1`, 0, 0, false, 15);
    }
    designSheet.activate();
    desighedText(designSheet, "✅  Done!", 10 + totalJobs, 6, 1, 1, "white", "Ink Free", 11, "green", true, "left");
}
function desighedText(sheet: ExcelScript.Worksheet, text: string, cursorrow: number, cursorcol: number, rows: number, cols: number, fillclr: string = "white", fontname: string = "Ink Free", fontsz: number = 10, fontclr: string = "black", fontbold: boolean = false, horalgn: string = "left", veralgn: string = "center") {
    const rng = sheet.getRangeByIndexes(cursorrow, cursorcol, rows, cols);
    rng.setValue(text);
    const fmt = rng.getFormat();
    fmt.getFill().setColor(fillclr);
    fmt.getFont().setName(fontname);
    fmt.getFont().setSize(fontsz);
    fmt.getFont().setColor(fontclr);
    fmt.getFont().setBold(fontbold);
    let halgn_ = ExcelScript.HorizontalAlignment.left;
    if (horalgn == "right")
        halgn_ = ExcelScript.HorizontalAlignment.right;
    else if (horalgn == "center")
        halgn_ = ExcelScript.HorizontalAlignment.center;
    let valgn_ = ExcelScript.VerticalAlignment.center;
    if (veralgn == "top")
        valgn_ = ExcelScript.VerticalAlignment.top;
    else if (veralgn == "bottom")
        valgn_ = ExcelScript.VerticalAlignment.bottom;
    fmt.setVerticalAlignment(valgn_);
    fmt.setHorizontalAlignment(halgn_);
}
function topHeads(sheet: ExcelScript.Worksheet, title: string, hor_start: number, hor_end: number, fillclr: string, fontsz: number, fontclr: string) {
    let horheader1 = sheet.getRange(convertToA1(1, hor_start) + ":" + convertToA1(1, hor_end));
    horheader1.merge();
    horheader1.setValue(title);
    horheader1.getFormat().getFont().setSize(fontsz);
    horheader1.getFormat().getFont().setColor(fontclr);
    horheader1.getFormat().getFill().setColor(fillclr);
    horheader1.getFormat().getFont().setName("Ink Free");
    horheader1.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    horheader1.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    const fmt = horheader1.getFormat();
    const borders = [ExcelScript.BorderIndex.edgeBottom];
    for (let b of borders) {
        fmt.getRangeBorder(b).setStyle(ExcelScript.BorderLineStyle.continuous);
        fmt.getRangeBorder(b).setWeight(3);
        fmt.getRangeBorder(b).setColor("#A6A6A6");
    }
}
function makeArrow(sheet: ExcelScript.Worksheet, direction: string, cursorrow: number, cursorcol: number, length: number) {
    let arrow = "⟶";
    let clr = "";
    let nrows = 1;
    let ncols = 1;
    if (direction == "left2right") {
        arrow = "⟶";
        clr = "gray";
        nrows = length;
        ncols = 1;
    }
    else if (direction == "top2bottom") {
        arrow = "↓";
        clr = "CornflowerBlue";
        nrows = 1;
        ncols = length;
    }
    else if (direction == "bottom2top") {
        arrow = "↑";
        clr = "tomato";
        nrows = 1;
        ncols = length;
    }
    const arrowrange = sheet.getRangeByIndexes(cursorrow, cursorcol, nrows, ncols);
    arrowrange.setValue(arrow);
    arrowrange.getFormat().getFont().setSize(9);
    arrowrange.getFormat().getFont().setColor(clr);
    arrowrange.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
}
function titleMaker(sheet: ExcelScript.Worksheet, title: string, titleUp: number = 0, titleRight: number = 2) {
    sheet.getFreezePanes().unfreeze();
    sheet.getFreezePanes().freezeRows(titleUp + 2);
    sheet.getFreezePanes().freezeColumns(1);
    sheet.getCell(titleUp, titleRight).setValue(title);
    const fmttitle = sheet.getCell(titleUp, titleRight).getFormat();
    fmttitle.setRowHeight(40);
    fmttitle.getFont().setSize(25);
    fmttitle.getFont().setBold(true);
    fmttitle.getFont().setColor("Maroon");
    fmttitle.getFont().setName("Ink free");
    fmttitle.setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
}
function getTorchCodeAnchorName(linkName: string): string {
    return linkName;
}
function getLayerCodeAnchor(layerIndex: number): string {
    return `code_layer_${layerIndex}`;
}
function goCode(sheet: ExcelScript.Worksheet, cursorrow: number, cursorcol: number, linkName: string, label: string = "</>") {
    if (cursorrow <= 0) {
        return;
    }
    const codeCell = sheet.getCell(cursorrow - 1, cursorcol);
    codeCell.setFormula(`=HYPERLINK("#Blueprint!torchcode", "${label}")`);
    codeCell.getFormat().getFont().setColor("gray");
    codeCell.getFormat().getFont().setSize(10);
}
function goTop(sheet: ExcelScript.Worksheet, cursorrow: number, cursorcol: number) {
    let formula_top = ``;
    formula_top = `=HYPERLINK("#X", "▲")`;
    sheet.getCell(cursorrow, cursorcol).setFormula(formula_top);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setColor("gray");
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setSize(11);
}
function goBlockLinks(sheet: ExcelScript.Worksheet, cursorrow: number, cursorcol: number, linkName: string, codeLabel: string = "</>") {
    goCode(sheet, cursorrow, cursorcol, linkName, codeLabel);
    goTop(sheet, cursorrow, cursorcol);
}
function genMultiSeriesChart(sheet: ExcelScript.Worksheet, title: string, datarow: number, datacol: number, r: number, c: number, chartLeft: number, chartTop: number, xAxisTitle?: string, yAxisTitle?: string) {
    if (r <= 1 || c <= 1) {
        return;
    }
    const xRange = sheet.getRangeByIndexes(datarow + 1, datacol, r - 1, 1);
    const dataRange = sheet.getRangeByIndexes(datarow, datacol + 1, r, c - 1);
    const chart = sheet.addChart(ExcelScript.ChartType.lineMarkers, dataRange);
    chart.setData(dataRange, ExcelScript.ChartSeriesBy.columns);
    chart.setPosition(convertToA1(chartTop, chartLeft), convertToA1(chartTop + 16, chartLeft + Math.max(12, c + 8)));
    chart.getTitle().setText(title);
    chart.getTitle().getFormat().getFont().setSize(18);
    chart.getTitle().getFormat().getFont().setBold(true);
    chart.getTitle().getFormat().getFont().setName("Ink Free");
    chart.getTitle().getFormat().getFont().setColor("Red");
    chart.getLegend().setVisible(true);
    chart.getLegend().setPosition(ExcelScript.ChartLegendPosition.right);
    const xAxis = chart.getAxes().getCategoryAxis();
    const yAxis = chart.getAxes().getValueAxis();
    xAxis.getTitle().setText(xAxisTitle);
    yAxis.getTitle().setText(yAxisTitle);
    xAxis.getTitle().getFormat().getFont().setSize(12);
    xAxis.getTitle().getFormat().getFont().setName("Ink Free");
    yAxis.getTitle().getFormat().getFont().setSize(12);
    yAxis.getTitle().getFormat().getFont().setName("Ink Free");
    const series = chart.getSeries();
    for (let i = 0; i < series.length; i++) {
        series[i].setXAxisValues(xRange);
        series[i].setMarkerStyle(ExcelScript.ChartMarkerStyle.circle);
        series[i].setMarkerSize(7);
        series[i].getFormat().getLine().setWeight(2);
    }
    chart.getFormat().getBorder().clear();
}
function genChart(sheet: ExcelScript.Worksheet, title: string, datarow: number, datacol: number, r: number, c: number, chartLeft: number, chartTop: number, xAxisTitle?: string, yAxisTitle?: string) {
    const range = sheet.getRangeByIndexes(datarow, datacol, r, c);
    const chart = sheet.addChart(ExcelScript.ChartType.xyscatterSmooth, range);
    chart.setPosition(convertToA1(chartTop, chartLeft), convertToA1(chartTop + 15, chartLeft + 20));
    chart.getTitle().setText(title);
    chart.getTitle().getFormat().getFont().setSize(20);
    chart.getTitle().getFormat().getFont().setBold(true);
    chart.getTitle().getFormat().getFont().setName("Ink Free");
    chart.getTitle().getFormat().getFont().setColor("Red");
    chart.getLegend().setVisible(false);
    const xAxis = chart.getAxes().getCategoryAxis();
    const yAxis = chart.getAxes().getValueAxis();
    xAxis.getTitle().setText(xAxisTitle);
    yAxis.getTitle().setText(yAxisTitle);
    xAxis.getTitle().getFormat().getFont().setSize(13);
    xAxis.getTitle().getFormat().getFont().setName("Ink Free");
    yAxis.getTitle().getFormat().getFont().setSize(13);
    yAxis.getTitle().getFormat().getFont().setName("Ink Free");
    const series = chart.getSeries()[0];
    series.setSmooth(true);
    series.setMarkerStyle(ExcelScript.ChartMarkerStyle.circle);
    series.setMarkerSize(10);
    series.getFormat().getLine().setColor("black");
    series.getFormat().getLine().setWeight(2);
    const points = series.getPoints();
    for (let i = 0; i < points.length; i++) {
        points[i].setMarkerStyle(ExcelScript.ChartMarkerStyle.circle);
        points[i].setMarkerSize(10);
        points[i].setMarkerBackgroundColor("#FFFF00");
        points[i].setMarkerForegroundColor("#000000");
    }
    chart.getFormat().getBorder().setColor("white");
}
function writeEquation(sheet: ExcelScript.Worksheet, header: string, cursorrow: number, cursorcol: number, algn?: string, input_var_name?: string) {
    let algn_ = ExcelScript.HorizontalAlignment.left;
    if (algn == "right")
        algn_ = ExcelScript.HorizontalAlignment.right;
    else if (algn == "center")
        algn_ = ExcelScript.HorizontalAlignment.center;
    sheet.getCell(cursorrow, cursorcol).setValue(header);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setSize(9);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setName("Ink Free");
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setBold(false);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setColor("red");
    sheet.getCell(cursorrow, cursorcol).getFormat().setHorizontalAlignment(algn_);
    sheet.getCell(cursorrow, cursorcol).getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
}
function makeKey(i: number, j: number): string {
    return `${i},${j}`;
}
function headerFormat(sheet: ExcelScript.Worksheet, header: string, cursorrow: number, cursorcol: number, algn?: string, fntcol: string = "black") {
    let algn_ = ExcelScript.HorizontalAlignment.left;
    if (algn == "right")
        algn_ = ExcelScript.HorizontalAlignment.right;
    else if (algn == "center")
        algn_ = ExcelScript.HorizontalAlignment.center;
    sheet.getCell(cursorrow, cursorcol).setValue(header);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setSize(11);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setBold(false);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setColor(fntcol);
    sheet.getCell(cursorrow, cursorcol).getFormat().setHorizontalAlignment(algn_);
    sheet.getCell(cursorrow, cursorcol).getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
}
function moveCursor(cursor: {
    row: number;
    col: number;
}, r: number, c: number) {
    cursor.row += r;
    cursor.col += c;
}
function updateName(workbook: ExcelScript.Workbook, cursor: {
    row: number;
    col: number;
    range: ExcelScript.Range;
}, namedVariable: string, sheet?: ExcelScript.Worksheet) {
    let temprange: ExcelScript.Range;
    temprange = cursor.range;
    let allNames = workbook.getNames();
    allNames.forEach(n => {
        if (n.getName() === namedVariable) {
            n.delete();
        }
    });
    if (sheet) {
        sheet.addNamedItem(namedVariable, temprange);
    }
    else {
        workbook.addNamedItem(namedVariable, temprange);
    }
}
function convertToA1(row: number, column: number): string {
    let columnLabel = "";
    column = column + 1;
    row = row + 1;
    while (column > 0) {
        column--;
        columnLabel = String.fromCharCode(65 + (column % 26)) + columnLabel;
        column = Math.floor(column / 26);
    }
    return `${columnLabel}${row}`;
}
function doImageFormatting(sheet: ExcelScript.Worksheet, cursorrow: number, cursorcol: number, number_of_rows: number, number_of_cols: number, lastRowOne: boolean = false) {
    const rowsForFormatting = lastRowOne ? number_of_rows - 1 : number_of_rows;
    let range = sheet.getRange(convertToA1(cursorrow, cursorcol) + ":" +
        convertToA1(cursorrow + rowsForFormatting - 1, cursorcol + number_of_cols - 1));
    range.clearAllConditionalFormats();
    const colorScale = range.addConditionalFormat(ExcelScript.ConditionalFormatType.colorScale);
    const cf = colorScale.getColorScale();
    cf.setCriteria({
        minimum: {
            type: ExcelScript.ConditionalFormatColorCriterionType.number,
            formula: "0",
            color: "dimgray"
        },
        maximum: {
            type: ExcelScript.ConditionalFormatColorCriterionType.number,
            formula: "1",
            color: "WhiteSmoke"
        }
    });
    const fmt = range.getFormat();
    const borders = [
        ExcelScript.BorderIndex.edgeTop,
        ExcelScript.BorderIndex.edgeBottom,
        ExcelScript.BorderIndex.edgeLeft,
        ExcelScript.BorderIndex.edgeRight,
    ];
    for (let b of borders) {
        fmt.getRangeBorder(b).setStyle(ExcelScript.BorderLineStyle.continuous);
        fmt.getRangeBorder(b).setWeight(2);
        fmt.getRangeBorder(b).setColor("Black");
    }
    range = sheet.getRange(convertToA1(cursorrow, cursorcol) + ":" + convertToA1(cursorrow + number_of_rows - 1, cursorcol + number_of_cols - 1));
    return range;
}
function doConditionalFormatting(sheet: ExcelScript.Worksheet, cursorrow: number, cursorcol: number, number_of_rows: number, number_of_cols: number, lowColor: string = "#2C7BB6", highColor: string = "#D7191C", vertEdge: boolean = true, horEdge: boolean = false, lastRowOne: boolean = false) {
    const rowsForFormatting = lastRowOne ? number_of_rows - 1 : number_of_rows;
    let range = sheet.getRange(convertToA1(cursorrow, cursorcol) + ":" +
        convertToA1(cursorrow + rowsForFormatting - 1, cursorcol + number_of_cols - 1));
    range.clearAllConditionalFormats();
    const colorScale = range.addConditionalFormat(ExcelScript.ConditionalFormatType.colorScale);
    const cf = colorScale.getColorScale();
    cf.setCriteria({
        minimum: {
            type: ExcelScript.ConditionalFormatColorCriterionType.lowestValue,
            color: lowColor
        },
        midpoint: {
            type: ExcelScript.ConditionalFormatColorCriterionType.percentile,
            formula: "50",
            color: "WhiteSmoke"
        },
        maximum: {
            type: ExcelScript.ConditionalFormatColorCriterionType.highestValue,
            color: highColor
        }
    });
    const fmt = range.getFormat();
    let borders = [
        ExcelScript.BorderIndex.edgeTop,
        ExcelScript.BorderIndex.edgeBottom,
        ExcelScript.BorderIndex.edgeLeft,
        ExcelScript.BorderIndex.edgeRight,
        ExcelScript.BorderIndex.insideVertical,
        ExcelScript.BorderIndex.insideHorizontal,
    ];
    if (horEdge) {
        borders = [
            ExcelScript.BorderIndex.edgeTop,
            ExcelScript.BorderIndex.edgeBottom,
            ExcelScript.BorderIndex.edgeLeft,
            ExcelScript.BorderIndex.edgeRight,
            ExcelScript.BorderIndex.insideHorizontal,
        ];
    }
    if (vertEdge) {
        borders = [
            ExcelScript.BorderIndex.edgeTop,
            ExcelScript.BorderIndex.edgeBottom,
            ExcelScript.BorderIndex.edgeLeft,
            ExcelScript.BorderIndex.edgeRight,
            ExcelScript.BorderIndex.insideVertical,
        ];
    }
    for (let b of borders) {
        fmt.getRangeBorder(b).setStyle(ExcelScript.BorderLineStyle.continuous);
        fmt.getRangeBorder(b).setColor("black");
    }
    range = sheet.getRange(convertToA1(cursorrow, cursorcol) + ":" + convertToA1(cursorrow + number_of_rows - 1, cursorcol + number_of_cols - 1));
    return range;
}
function doFormatting(sheet: ExcelScript.Worksheet, cursorrow: number, cursorcol: number, number_of_rows: number, number_of_cols: number, color: string) {
    const range = sheet.getRange(convertToA1(cursorrow, cursorcol) + ":" + convertToA1(cursorrow + number_of_rows - 1, cursorcol + number_of_cols - 1));
    const fmt = range.getFormat();
    const borders = [
        ExcelScript.BorderIndex.edgeTop,
        ExcelScript.BorderIndex.edgeBottom,
        ExcelScript.BorderIndex.edgeLeft,
        ExcelScript.BorderIndex.edgeRight,
        ExcelScript.BorderIndex.insideVertical,
        ExcelScript.BorderIndex.insideHorizontal,
    ];
    for (let b of borders) {
        fmt.getRangeBorder(b).setStyle(ExcelScript.BorderLineStyle.continuous);
        fmt.getRangeBorder(b).setColor("gray");
    }
    fmt.getFill().setColor(color);
    return range;
}
function doFormatting4NamedItems(sheet: ExcelScript.Worksheet, cursorrow: number, cursorcol: number, number_of_rows: number, number_of_cols: number, color: string) {
    const range = sheet.getRange(convertToA1(cursorrow, cursorcol) + ":" + convertToA1(cursorrow + number_of_rows - 1, cursorcol + number_of_cols - 1));
    const fmt = range.getFormat();
    const borders = [
        ExcelScript.BorderIndex.edgeTop,
        ExcelScript.BorderIndex.edgeBottom,
        ExcelScript.BorderIndex.edgeLeft,
        ExcelScript.BorderIndex.edgeRight,
        ExcelScript.BorderIndex.insideVertical,
        ExcelScript.BorderIndex.insideHorizontal,
    ];
    for (let b of borders) {
        fmt.getRangeBorder(b).setStyle(ExcelScript.BorderLineStyle.continuous);
        fmt.getRangeBorder(b).setColor("gray");
    }
    fmt.getFill().setColor(color);
    return range;
}
function embedViaLookUp(tokens: string, sheet: ExcelScript.Worksheet, inputRange: ExcelScript.Range) {
    const dim = 3;
    const lookup = lookUp();
    const embedMat: number[][] = Array.from({ length: dim }, () => []);
    for (const token of tokens) {
        const vec = lookup[token];
        if (vec) {
            for (let i = 0; i < dim; i++)
                embedMat[i].push(vec[i]);
        }
        else {
            for (let i = 0; i < dim; i++)
                embedMat[i].push(0);
        }
    }
    return embedMat;
}
function displayImages(workbook: ExcelScript.Workbook, sheet: ExcelScript.Worksheet, header: string, cursorrow: number, cursorcol: number, dataset: string, n_class: number, n_smpls: number, n_batch: number, number_of_rows: number, number_of_cols: number, ones: boolean = false, maxDim: number = 0) {
    const tempcursorcol = cursorcol;
    cursorcol = 4;
    sheet.getCell(cursorrow, cursorcol).setValue(header);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setSize(12);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setBold(true);
    sheet.getCell(cursorrow, cursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    cursorrow += 1;
    let rand_cls = 0;
    let rand_smpls = 0;
    let formula_ = `=RANDBETWEEN(0,${n_class - 1})`;
    let formula_s = `=RANDBETWEEN(0,${n_smpls - 1})`;
    let shiftright = 1;
    let shiftdown = 0;
    let gap = 0;
    let n_rows_withones = 0;
    for (let b = 0; b < n_batch; b++) {
        rand_cls = Math.floor(Math.random() * n_class);
        rand_smpls = Math.floor(Math.random() * n_smpls);
        shiftdown = 0;
        sheet.getCell(cursorrow + shiftdown, shiftright + cursorcol + (number_of_cols + gap) * b).setValue("[" + String(b + 1) + "]");
        sheet.getCell(cursorrow + shiftdown, shiftright + cursorcol + (number_of_cols + gap) * b).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
        shiftdown = 1;
        sheet.getCell(cursorrow + shiftdown, shiftright + cursorcol + (number_of_cols + gap) * b - 1).setValue("class:");
        sheet.getCell(cursorrow + shiftdown, shiftright + cursorcol + (number_of_cols + gap) * b).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
        sheet.getCell(cursorrow + shiftdown, shiftright + cursorcol + (number_of_cols + gap) * b).setValue(rand_cls);
        sheet.getCell(cursorrow + shiftdown, shiftright + cursorcol + (number_of_cols + gap) * b).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
        shiftdown = 2;
        sheet.getCell(cursorrow + shiftdown, shiftright + cursorcol + (number_of_cols + gap) * b - 1).setValue("sample:");
        sheet.getCell(cursorrow + shiftdown, shiftright + cursorcol + (number_of_cols + gap) * b - 1).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
        sheet.getCell(cursorrow + shiftdown, shiftright + cursorcol + (number_of_cols + gap) * b).setValue(rand_smpls);
        sheet.getCell(cursorrow + shiftdown, shiftright + cursorcol + (number_of_cols + gap) * b).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
        let formula_im = ``;
        formula_im = `=OFFSET(${dataset}!C4,${number_of_rows + 2}*${convertToA1(cursorrow + 1, shiftright + cursorcol + (number_of_cols + gap) * b)},${number_of_cols + 2}*${convertToA1(cursorrow + 1 + 1, shiftright + cursorcol + (number_of_cols + gap) * b)}, ${number_of_rows}, ${number_of_cols})`;
        formula_im = ones ? `=VSTACK(${formula_im.substring(1)}, SEQUENCE(1,${number_of_cols})*0+1)` : formula_im;
        n_rows_withones = ones ? 1 : 0;
        shiftdown = 3;
        sheet
            .getCell(cursorrow + shiftdown, (shiftright - 1) + cursorcol + (number_of_cols + gap) * b)
            .setFormula(formula_im);
        const range = doImageFormatting(sheet, cursorrow + shiftdown, (shiftright - 1) + cursorcol + (number_of_cols + gap) * b, number_of_rows + 1 * n_rows_withones, number_of_cols, ones);
        const formula_ohe = `=TRANSPOSE(--(SEQUENCE(1, ${n_class}) = 1+${convertToA1(cursorrow + 1, shiftright + cursorcol + (number_of_cols + gap) * b)}))`;
        shiftdown += number_of_rows + 2;
        sheet
            .getCell(cursorrow + shiftdown, cursorcol + b)
            .setFormula(formula_ohe);
    }
    cursorrow += 3;
    const range = sheet.getRange(convertToA1(cursorrow, cursorcol) + ":" + convertToA1(cursorrow + number_of_rows - 1, cursorcol + (number_of_cols + gap) * n_batch - (gap + 1)));
    sheet.getCell(cursorrow + number_of_rows + 1, cursorcol).setValue("target = OHE");
    sheet.getCell(cursorrow + number_of_rows + 1, cursorcol).getFormat().getFont().setSize(12);
    sheet.getCell(cursorrow + number_of_rows + 1, cursorcol).getFormat().getFont().setBold(true);
    sheet.getCell(cursorrow + number_of_rows + 1, cursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    const range_ohe = doConditionalFormatting(sheet, cursorrow + number_of_rows + 2, cursorcol, n_class, n_batch, "WhiteSmoke", "Red", true, false, false);
    sheet.addNamedItem("ohe", range_ohe);
    cursorcol = tempcursorcol;
    cursorrow = cursorrow + number_of_rows + 1 + 1 * n_rows_withones;
    cursorrow += n_class + 2;
    return { row: cursorrow, col: cursorcol, range };
}
function displayImages2(workbook: ExcelScript.Workbook, sheet: ExcelScript.Worksheet, header: string, cursorrow: number, cursorcol: number, dataset: string, n_class: number, n_smpls: number, n_batch: number, number_of_rows: number, number_of_cols: number, ones: boolean = false, hasLoss: boolean = false, maxDim: number = 0) {
    const tempcursorcol = cursorcol;
    cursorcol = 4;
    sheet.getCell(cursorrow, cursorcol).setValue(header);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setSize(12);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setBold(true);
    sheet.getCell(cursorrow, cursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    cursorrow += 1;
    let rand_cls = 0;
    let rand_smpls = 0;
    let formula_ = `=RANDBETWEEN(0,${n_class - 1})`;
    let formula_s = `=RANDBETWEEN(0,${n_smpls - 1})`;
    let shiftright = 1;
    let shiftdown = 0;
    let gap = 0;
    let n_rows_withones = 0;
    for (let b = 0; b < n_batch; b++) {
        rand_cls = Math.floor(Math.random() * n_class);
        rand_smpls = Math.floor(Math.random() * n_smpls);
        shiftdown = 0;
        sheet.getCell(cursorrow + shiftdown, shiftright + cursorcol + (number_of_cols + gap) * b).setValue("Image # " + String(b + 1));
        sheet.getCell(cursorrow + shiftdown, shiftright + cursorcol + (number_of_cols + gap) * b).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
        sheet.getCell(cursorrow + shiftdown, shiftright + cursorcol + (number_of_cols + gap) * b).getFormat().getFont().setColor("blue");
        shiftdown = 1;
        sheet.getCell(cursorrow + shiftdown, shiftright + cursorcol + (number_of_cols + gap) * b - 1).setValue("class:");
        sheet.getCell(cursorrow + shiftdown, shiftright + cursorcol + (number_of_cols + gap) * b).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
        sheet.getCell(cursorrow + shiftdown, shiftright + cursorcol + (number_of_cols + gap) * b).setValue(rand_cls);
        sheet.getCell(cursorrow + shiftdown, shiftright + cursorcol + (number_of_cols + gap) * b).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
        shiftdown = 2;
        sheet.getCell(cursorrow + shiftdown, shiftright + cursorcol + (number_of_cols + gap) * b - 1).setValue("sample:");
        sheet.getCell(cursorrow + shiftdown, shiftright + cursorcol + (number_of_cols + gap) * b - 1).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
        sheet.getCell(cursorrow + shiftdown, shiftright + cursorcol + (number_of_cols + gap) * b).setValue(rand_smpls);
        sheet.getCell(cursorrow + shiftdown, shiftright + cursorcol + (number_of_cols + gap) * b).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
        let formula_im = ``;
        formula_im = `=OFFSET(${dataset}!C4,${number_of_rows + 2}*${convertToA1(cursorrow + 1, shiftright + cursorcol + (number_of_cols + gap) * b)},${number_of_cols + 2}*${convertToA1(cursorrow + 1 + 1, shiftright + cursorcol + (number_of_cols + gap) * b)}, ${number_of_rows}, ${number_of_cols})`;
        formula_im = ones ? `=VSTACK(${formula_im.substring(1)}, SEQUENCE(1,${number_of_cols})*0+1)` : formula_im;
        n_rows_withones = ones ? 1 : 0;
        shiftdown = 3;
        sheet
            .getCell(cursorrow + shiftdown, (shiftright - 1) + cursorcol + (number_of_cols + gap) * b)
            .setFormula(formula_im);
        const range = doImageFormatting(sheet, cursorrow + shiftdown, (shiftright - 1) + cursorcol + (number_of_cols + gap) * b, number_of_rows + 1 * n_rows_withones, number_of_cols, ones);
        if (hasLoss) {
            const formula_ohe = `=TRANSPOSE(--(SEQUENCE(1, ${n_class}) = 1+${convertToA1(cursorrow + 1, shiftright + cursorcol + (number_of_cols + gap) * b)}))`;
            shiftdown += number_of_rows + 2;
            sheet
                .getCell(cursorrow + shiftdown, cursorcol + b)
                .setFormula(formula_ohe);
        }
    }
    cursorrow += 3;
    const range = sheet.getRange(convertToA1(cursorrow, cursorcol) + ":" + convertToA1(cursorrow + number_of_rows - 1, cursorcol + (number_of_cols + gap) * n_batch - (gap + 1)));
    if (hasLoss) {
        sheet.getCell(cursorrow + number_of_rows + 1, cursorcol).setValue(`target = OHE(${n_class})`);
        sheet.getCell(cursorrow + number_of_rows + 1, cursorcol).getFormat().getFont().setSize(12);
        sheet.getCell(cursorrow + number_of_rows + 1, cursorcol).getFormat().getFont().setBold(true);
        sheet.getCell(cursorrow + number_of_rows + 1, cursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
        const range_ohe = doConditionalFormatting(sheet, cursorrow + number_of_rows + 2, cursorcol, n_class, n_batch, "WhiteSmoke", "Red", true, false, false);
        sheet.addNamedItem("ohe", range_ohe);
        cursorrow += n_class;
    }
    cursorcol = tempcursorcol;
    cursorrow = cursorrow + number_of_rows + 1 + 1 * n_rows_withones;
    cursorrow += 3;
    return { row: cursorrow, col: cursorcol, range };
}
function displayRandImages(workbook: ExcelScript.Workbook, sheet: ExcelScript.Worksheet, header: string, cursorrow: number, cursorcol: number, dataset: string, n_class: number, n_smpls: number, n_batch: number, number_of_rows: number, number_of_cols: number, ones: boolean = false, hasLoss: boolean = false, maxDim: number = 0) {
    const tempcursorcol = cursorcol;
    cursorcol = 4;
    sheet.getCell(cursorrow, cursorcol).setValue(header);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setSize(12);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setBold(true);
    sheet.getCell(cursorrow, cursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    cursorrow += 1;
    let rand_cls = 0;
    let rand_smpls = 0;
    let formula_ = `=RANDBETWEEN(0,${n_class - 1})`;
    let formula_s = `=RANDBETWEEN(0,${n_smpls - 1})`;
    let shiftright = 1;
    let shiftdown = 0;
    let gap = 0;
    let n_rows_withones = 0;
    for (let b = 0; b < n_batch; b++) {
        rand_cls = Math.floor(Math.random() * n_class);
        rand_smpls = Math.floor(Math.random() * n_smpls);
        shiftdown = 0;
        sheet.getCell(cursorrow + shiftdown, shiftright + cursorcol + (number_of_cols + gap) * b).setValue("Image # " + String(b + 1));
        sheet.getCell(cursorrow + shiftdown, shiftright + cursorcol + (number_of_cols + gap) * b).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
        sheet.getCell(cursorrow + shiftdown, shiftright + cursorcol + (number_of_cols + gap) * b).getFormat().getFont().setColor("blue");
        let formula_im = ``;
        formula_im = `=RANDARRAY(${number_of_rows}, ${number_of_cols}, 0,1)`;
        formula_im = ones ? `=VSTACK(${formula_im.substring(1)}, SEQUENCE(1,${number_of_cols})*0+1)` : formula_im;
        n_rows_withones = ones ? 1 : 0;
        shiftdown = 1;
        sheet
            .getCell(cursorrow + shiftdown, (shiftright - 1) + cursorcol + (number_of_cols + gap) * b)
            .setFormula(formula_im);
        const range = doImageFormatting(sheet, cursorrow + shiftdown, (shiftright - 1) + cursorcol + (number_of_cols + gap) * b, number_of_rows + 1 * n_rows_withones, number_of_cols, ones);
    }
    cursorrow += 1;
    const range = sheet.getRange(convertToA1(cursorrow, cursorcol) + ":" + convertToA1(cursorrow + number_of_rows - 1, cursorcol + (number_of_cols + gap) * n_batch - (gap + 1)));
    if (hasLoss) {
        let tempcursorrow = cursorrow;
        let tempcursorcol = cursorcol;
        tempcursorrow = tempcursorrow + number_of_rows + 1 * n_rows_withones;
        tempcursorrow++;
        sheet.getCell(tempcursorrow, tempcursorcol).setValue(`target = OHE(${n_class})`);
        sheet.getCell(tempcursorrow, tempcursorcol).getFormat().getFont().setSize(12);
        sheet.getCell(tempcursorrow, tempcursorcol).getFormat().getFont().setBold(true);
        sheet.getCell(tempcursorrow, tempcursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
        sheet.getCell(tempcursorrow + 1, tempcursorcol - 1).setValue("Class ⟶  ");
        sheet.getCell(tempcursorrow + 1, tempcursorcol - 1).getFormat().getFont().setBold(false);
        sheet.getCell(tempcursorrow + 1, tempcursorcol - 1).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
        tempcursorrow++;
        for (let b = 0; b < n_batch; b++) {
            let rand_cls = Math.floor(Math.random() * n_class);
            dropDownSequence(sheet, tempcursorrow, tempcursorcol + b, n_class, rand_cls);
            sheet.getCell(tempcursorrow, tempcursorcol + b)
                .getFormat()
                .getFill()
                .setColor("#f1f4ed");
            const formula_ohe = `=TRANSPOSE(--(SEQUENCE(1, ${n_class}) = 1+${convertToA1(tempcursorrow, tempcursorcol + b)}))`;
            sheet
                .getCell(tempcursorrow + 1, tempcursorcol + b)
                .setFormula(formula_ohe);
        }
        const range_ohe = doConditionalFormatting(sheet, tempcursorrow + 1, tempcursorcol, n_class, n_batch, "WhiteSmoke", "Red", true, false, false);
        sheet.addNamedItem("ohe", range_ohe);
        cursorrow += n_class + 3;
    }
    cursorcol = tempcursorcol;
    cursorrow = cursorrow + number_of_rows + 1 + 1 * n_rows_withones;
    cursorrow += 3;
    return { row: cursorrow, col: cursorcol, range };
}
function displayRandImages2(workbook: ExcelScript.Workbook, sheet: ExcelScript.Worksheet, header: string, cursorrow: number, cursorcol: number, dataset: string, n_class: number, n_smpls: number, n_batch: number, number_of_rows: number, number_of_cols: number, ones: boolean = false, hasLoss: boolean = false, maxDim: number = 0, randorseq: string = "random", seed: number = 1) {
    const tempcursorcol = cursorcol;
    cursorcol = 4;
    sheet.getCell(cursorrow, cursorcol).setValue(header);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setSize(12);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setBold(true);
    sheet.getCell(cursorrow, cursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    cursorrow += 1;
    let rand_cls = 0;
    let rand_smpls = 0;
    let formula_ = `=RANDBETWEEN(0,${n_class - 1})`;
    let formula_s = `=RANDBETWEEN(0,${n_smpls - 1})`;
    let shiftright = 1;
    let shiftdown = 0;
    let gap = 0;
    let n_rows_withones = 0;
    for (let b = 0; b < n_batch; b++) {
        rand_cls = Math.floor(Math.random() * n_class);
        rand_smpls = Math.floor(Math.random() * n_smpls);
        shiftdown = 0;
        sheet.getCell(cursorrow + shiftdown, shiftright + cursorcol + (number_of_cols + gap) * b).setValue("Image # " + String(b + 1));
        sheet.getCell(cursorrow + shiftdown, shiftright + cursorcol + (number_of_cols + gap) * b).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
        sheet.getCell(cursorrow + shiftdown, shiftright + cursorcol + (number_of_cols + gap) * b).getFormat().getFont().setColor("blue");
        if (randorseq === "seeded") {
            seed += 100;
            let state = seed >>> 0;
            const rand = () => {
                state = (1664525 * state + 1013904223) >>> 0;
                return state / 4294967296;
            };
            let data: number[][] = [];
            for (let i = 0; i < number_of_rows; i++) {
                let row: number[] = [];
                for (let j = 0; j < number_of_cols; j++) {
                    row.push(rand());
                }
                data.push(row);
            }
            if (ones) {
                data.push(new Array(number_of_cols).fill(1));
            }
            shiftdown = 1;
            sheet.getRangeByIndexes(cursorrow + shiftdown, (shiftright - 1) + cursorcol + (number_of_cols + gap) * b, number_of_rows + n_rows_withones, number_of_cols).setValues(data);
        }
        else {
            let formula_im = ``;
            formula_im = `=RANDARRAY(${number_of_rows}, ${number_of_cols}, 0,1)`;
            formula_im = ones ? `=VSTACK(${formula_im.substring(1)}, SEQUENCE(1,${number_of_cols})*0+1)` : formula_im;
            n_rows_withones = ones ? 1 : 0;
            shiftdown = 1;
            sheet.getCell(cursorrow + shiftdown, (shiftright - 1) + cursorcol + (number_of_cols + gap) * b).setFormula(formula_im);
        }
        const range = doImageFormatting(sheet, cursorrow + shiftdown, (shiftright - 1) + cursorcol + (number_of_cols + gap) * b, number_of_rows + 1 * n_rows_withones, number_of_cols, ones);
    }
    cursorrow += 1;
    const range = sheet.getRange(convertToA1(cursorrow, cursorcol) + ":" + convertToA1(cursorrow + number_of_rows - 1, cursorcol + (number_of_cols + gap) * n_batch - (gap + 1)));
    if (hasLoss) {
        let tempcursorrow = cursorrow;
        let tempcursorcol = cursorcol;
        tempcursorrow = tempcursorrow + number_of_rows + 1 * n_rows_withones;
        tempcursorrow++;
        sheet.getCell(tempcursorrow, tempcursorcol).setValue(`target = OHE(${n_class})`);
        sheet.getCell(tempcursorrow, tempcursorcol).getFormat().getFont().setSize(12);
        sheet.getCell(tempcursorrow, tempcursorcol).getFormat().getFont().setBold(true);
        sheet.getCell(tempcursorrow, tempcursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
        sheet.getCell(tempcursorrow + 1, tempcursorcol - 1).setValue("Class ⟶  ");
        sheet.getCell(tempcursorrow + 1, tempcursorcol - 1).getFormat().getFont().setBold(false);
        sheet.getCell(tempcursorrow + 1, tempcursorcol - 1).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
        tempcursorrow++;
        for (let b = 0; b < n_batch; b++) {
            let rand_cls = Math.floor(Math.random() * n_class);
            dropDownSequence(sheet, tempcursorrow, tempcursorcol + b, n_class, rand_cls);
            sheet.getCell(tempcursorrow, tempcursorcol + b).getFormat().getFill().setColor("#f1f4ed");
            const formula_ohe = `=TRANSPOSE(--(SEQUENCE(1, ${n_class}) = 1+${convertToA1(tempcursorrow, tempcursorcol + b)}))`;
            sheet
                .getCell(tempcursorrow + 1, tempcursorcol + b)
                .setFormula(formula_ohe);
        }
        const range_ohe = doConditionalFormatting(sheet, tempcursorrow + 1, tempcursorcol, n_class, n_batch, "WhiteSmoke", "Red", true, false, false);
        sheet.addNamedItem("ohe", range_ohe);
        cursorrow += n_class + 3;
    }
    cursorcol = tempcursorcol;
    cursorrow = cursorrow + number_of_rows + 1 + 1 * n_rows_withones;
    cursorrow += 1;
    return { row: cursorrow, col: cursorcol, range };
}
function OHE(sheet: ExcelScript.Worksheet, header: string, cursorrow: number, cursorcol: number, dataset: string, n_class: number, n_smpls: number, n_batch: number, number_of_rows: number, number_of_cols: number, ones: boolean = false) {
}
function drawMatrix(sheet: ExcelScript.Worksheet, header: string, cursorrow: number, cursorcol: number, number_of_rows: number, number_of_cols: number, randorseq: string, in_range?: ExcelScript.Range) {
    sheet.getCell(cursorrow, cursorcol).setValue(header);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setSize(12);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setBold(true);
    sheet.getCell(cursorrow, cursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    cursorrow += 1;
    let formula_ = ``;
    let color = "#daf2d1";
    if (randorseq == "random") {
        formula_ = `=RANDARRAY(${number_of_rows}, ${number_of_cols}, -1, 1)`;
        color = "yellow";
    }
    else if (randorseq == "seq") {
        formula_ = `=SEQUENCE(${number_of_rows}, ${number_of_cols}, 0)/(${number_of_rows}*${number_of_cols})-0.5`;
        color = "white";
    }
    else if (randorseq == "other") {
        formula_ = `=${in_range.getAddress()}`;
    }
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range = doConditionalFormatting(sheet, cursorrow, cursorcol, number_of_rows, number_of_cols, "orange", "yellow", false, true);
    cursorrow += number_of_rows + 1;
    return { row: cursorrow, col: cursorcol, range };
}
function drawMatrix2(sheet: ExcelScript.Worksheet, header: string, cursorrow: number, cursorcol: number, number_of_rows: number, number_of_cols: number, randorseq: string, in_range?: ExcelScript.Range, ones: boolean = false) {
    sheet.getCell(cursorrow, cursorcol).setValue(header);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setSize(12);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setBold(true);
    sheet.getCell(cursorrow, cursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    cursorrow += 1;
    let formula_ = ``;
    let color = "#daf2d1";
    if (randorseq == "random") {
        formula_ = `=RANDARRAY(${number_of_rows}, ${number_of_cols}, -1, 1)`;
        color = "yellow";
    }
    else if (randorseq == "seq") {
        formula_ = `=SEQUENCE(${number_of_rows}, ${number_of_cols}, 0)`;
        color = "white";
    }
    else if (randorseq == "other") {
        formula_ = `=${in_range.getAddress()}`;
    }
    let n_rows_withones = 1;
    if (ones) {
        formula_ = `=VSTACK(WRAPCOLS(TOCOL(${in_range.getAddress()},1,TRUE),${number_of_rows}*${number_of_cols}), SEQUENCE(1,${number_of_cols})*0+1)`;
        n_rows_withones = 1;
    }
    else {
        formula_ = `=WRAPCOLS(TOCOL(${in_range.getAddress()},1,TRUE),${number_of_rows}*${number_of_cols})`;
        n_rows_withones = 0;
    }
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range = doConditionalFormatting(sheet, cursorrow, cursorcol, number_of_rows * number_of_cols + 1 * n_rows_withones, number_of_cols, "orange", "yellow", false, true);
    cursorrow += number_of_rows + number_of_rows * number_of_cols;
    return { row: cursorrow, col: cursorcol, range };
}
function drawMatrix3(sheet: ExcelScript.Worksheet, header: string, cursorrow: number, cursorcol: number, number_of_rows: number, number_of_cols: number, randorseq: string, in_range?: ExcelScript.Range, seed: number = 1, min: number = -1, max: number = 1) {
    cursorrow++;
    headerFormat(sheet, header, cursorrow, cursorcol + number_of_cols - 1, "right");
    cursorrow++;
    let color = "#daf2d1";
    if (randorseq === "random") {
        const formula_ = `=RANDARRAY(${number_of_rows}, ${number_of_cols}, -1, 1)`;
        sheet.getCell(cursorrow, cursorcol).setFormula(formula_);
        color = "yellow";
    }
    else if (randorseq === "seq") {
        const formula_ = `=SEQUENCE(${number_of_rows}, ${number_of_cols}, 0)
             /(${number_of_rows}*${number_of_cols}) - 0.5`;
        sheet.getCell(cursorrow, cursorcol).setFormula(formula_);
        color = "white";
    }
    else if (randorseq === "other" && in_range) {
        sheet.getCell(cursorrow, cursorcol)
            .setFormula(`=${in_range.getAddress()}`);
    }
    else if (randorseq === "seeded") {
        let state = seed >>> 0;
        const rand = () => {
            state = (1664525 * state + 1013904223) >>> 0;
            return state / 4294967296;
        };
        let data: number[][] = [];
        for (let i = 0; i < number_of_rows; i++) {
            let row: number[] = [];
            for (let j = 0; j < number_of_cols; j++) {
                const v = rand() * (max - min) + min;
                row.push(v);
            }
            data.push(row);
        }
        sheet
            .getRangeByIndexes(cursorrow, cursorcol, number_of_rows, number_of_cols)
            .setValues(data);
        color = "orange";
    }
    const range = doConditionalFormatting(sheet, cursorrow, cursorcol, number_of_rows, number_of_cols, "orange", "yellow", false, true);
    cursorrow += number_of_rows + 1;
    return { row: cursorrow, col: cursorcol, range };
}
function ABCto123(sheet: ExcelScript.Worksheet, header: string, cursorrow: number, cursorcol: number, in_range: ExcelScript.Range, number_of_rows: number, number_of_cols: number) {
    sheet.getCell(cursorrow, cursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    let formula_ = ``;
    formula_ = `=CODE(LOWER(${in_range.getAddress()}))- CODE("a")`;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range = doFormatting(sheet, cursorrow, cursorcol, number_of_rows, number_of_cols, "#d5a6bd");
    cursorrow += number_of_rows + 1;
    return { row: cursorrow, col: cursorcol, range };
}
function Flattening(sheet: ExcelScript.Worksheet, header: string, cursorrow: number, cursorcol: number, in_range: ExcelScript.Range, number_of_rows: number, number_of_cols: number, n_batch: number, ones: boolean = true, input_var_name?: string) {
    headerFormat(sheet, header, cursorrow, cursorcol);
    cursorrow += 2;
    let formula_ = ``;
    let n_rows_withones = 1;
    if (ones) {
        formula_ = `=VSTACK(WRAPCOLS(TOCOL(${input_var_name},1,TRUE),${number_of_rows}*${number_of_cols}), SEQUENCE(1,${n_batch})*0+1)`;
        n_rows_withones = 1;
    }
    else {
        formula_ = `=WRAPCOLS(TOCOL(${input_var_name},1,TRUE),${number_of_rows}*${number_of_cols})`;
        n_rows_withones = 0;
    }
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range = doConditionalFormatting(sheet, cursorrow, cursorcol, number_of_rows * number_of_cols + 1 * n_rows_withones, n_batch, "dimgray", "WhiteSmoke", true, false, ones);
    cursorrow += number_of_rows * number_of_cols + 1;
    return { row: cursorrow, col: cursorcol, range };
}
function PassIt(sheet: ExcelScript.Worksheet, header: string, cursorrow: number, cursorcol: number, in_range: ExcelScript.Range, number_of_rows: number, number_of_cols: number, ones: boolean = true) {
    sheet.getCell(cursorrow, cursorcol).setValue(header);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setSize(12);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setBold(true);
    sheet.getCell(cursorrow, cursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    cursorrow += 1;
    let formula_ = ``;
    formula_ = `=${in_range.getAddress()}`;
    formula_ = ones ? `=VSTACK(${formula_.substring(1)}, SEQUENCE(1,${number_of_cols})*0+1)` : formula_;
    let n_rows_withones = ones ? 1 : 0;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range = doConditionalFormatting(sheet, cursorrow, cursorcol, number_of_rows + 1 * n_rows_withones, number_of_cols, "CornflowerBlue", "Tomato", true, false, ones);
    cursorrow += number_of_rows + 1 * n_rows_withones;
    return { row: cursorrow, col: cursorcol, range };
}
function PE(sheet: ExcelScript.Worksheet, header: string, cursorrow: number, cursorcol: number, in_range: ExcelScript.Range, number_of_rows: number, number_of_cols: number, ones: boolean = true, input_var_name?: string) {
    let d_model = number_of_rows;
    sheet.getCell(cursorrow, cursorcol).setValue(header);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setSize(12);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setBold(true);
    sheet.getCell(cursorrow, cursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    cursorrow += 1;
    cursorrow += 1;
    let formula_ = ``;
    let pos_row = 0;
    let pos_col = 0;
    let i_row = 0;
    let i_col = 0;
    let formula_pos = `=SEQUENCE(1,${number_of_cols},0)`;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_pos);
    pos_row = cursorrow;
    pos_col = cursorcol;
    const range_pos = doConditionalFormatting(sheet, cursorrow, cursorcol, 1, number_of_cols, "WhiteSmoke", "WhiteSmoke", true, true, false);
    cursorrow += 3;
    let formula_i = `=SEQUENCE(${number_of_rows},1,0)`;
    cursorcol -= 2;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_i);
    i_row = cursorrow;
    i_col = cursorcol;
    const range_i = doConditionalFormatting(sheet, cursorrow, cursorcol, number_of_rows, 1, "WhiteSmoke", "WhiteSmoke", false, true, false);
    cursorcol += 2;
    formula_ = `=IF(MOD(${convertToA1(i_row, i_col)}#,2)=0, SIN(${convertToA1(pos_row, pos_col)}#/POWER(10000, ${convertToA1(i_row, i_col)}# / ${number_of_rows})), COS(${convertToA1(pos_row, pos_col)}#/POWER(10000, ${convertToA1(i_row, i_col)}# / ${number_of_rows})))`;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const perow = cursorrow;
    const pecol = cursorcol;
    const range_pe = doConditionalFormatting(sheet, cursorrow, cursorcol, number_of_rows, number_of_cols, "CornflowerBlue", "Tomato", true, false, false);
    cursorrow += number_of_rows;
    let formula_pex = ``;
    formula_pex = `=${input_var_name}+${convertToA1(perow, pecol)}#`;
    formula_pex = ones ? `=VSTACK(${formula_pex.substring(1)}, SEQUENCE(1,${number_of_cols})*0+1)` : formula_;
    let n_rows_withones = ones ? 1 : 0;
    cursorrow += 2;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_pex);
    const range = doConditionalFormatting(sheet, cursorrow, cursorcol, number_of_rows + 1 * n_rows_withones, number_of_cols, "CornflowerBlue", "Tomato", true, false, ones);
    cursorrow += number_of_rows + 1 * n_rows_withones;
    return { row: cursorrow, col: cursorcol, range };
}
function ReLU(sheet: ExcelScript.Worksheet, header: string, cursorrow: number, cursorcol: number, in_range: ExcelScript.Range, number_of_rows: number, number_of_cols: number, ones: boolean = true, input_var_name?: string) {
    headerFormat(sheet, header, cursorrow, cursorcol);
    cursorrow += 1;
    const _eq = `✏️ ${header.replace(/=.*/, "")} = max(${input_var_name} , 0)`;
    writeEquation(sheet, _eq, cursorrow, cursorcol, "left");
    cursorrow += 1;
    let formula_ = ``;
    formula_ = `=IF(${in_range.getAddress()}>0, ${in_range.getAddress()}, 0)`;
    formula_ = ones ? `=VSTACK(${formula_.substring(1)}, SEQUENCE(1,${number_of_cols})*0+1)` : formula_;
    let n_rows_withones = ones ? 1 : 0;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range = doConditionalFormatting(sheet, cursorrow, cursorcol, number_of_rows + 1 * n_rows_withones, number_of_cols, "CornflowerBlue", "Tomato", true, false, ones);
    cursorrow += number_of_rows + 1 * n_rows_withones;
    return { row: cursorrow, col: cursorcol, range };
}
function TanH(sheet: ExcelScript.Worksheet, header: string, cursorrow: number, cursorcol: number, in_range: ExcelScript.Range, number_of_rows: number, number_of_cols: number, ones: boolean = true, input_var_name?: string) {
    headerFormat(sheet, header, cursorrow, cursorcol);
    cursorrow += 1;
    const _eq = `✏️ 𝑒 = exp(${input_var_name})`;
    writeEquation(sheet, _eq, cursorrow, cursorcol, "left");
    cursorrow += 1;
    let formula_ = ``;
    formula_ = `=EXP(${in_range.getAddress()})`;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const rangeexp = doConditionalFormatting(sheet, cursorrow, cursorcol, number_of_rows, number_of_cols, "ForestGreen", "PaleGreen");
    cursorrow += number_of_rows;
    const _eq2 = `✏️ 𝑒⁻ = exp(−${input_var_name})`;
    writeEquation(sheet, _eq2, cursorrow, cursorcol, "left");
    cursorrow += 1;
    formula_ = `=EXP(-${in_range.getAddress()})`;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const rangeexp_neg = doConditionalFormatting(sheet, cursorrow, cursorcol, number_of_rows, number_of_cols, "ForestGreen", "PaleGreen");
    cursorrow += number_of_rows;
    const _eq3 = `✏️ ${header.replace(/=.*/, "")} = (𝑒-𝑒⁻)∕(𝑒+𝑒⁻)`;
    writeEquation(sheet, _eq3, cursorrow, cursorcol, "left");
    cursorrow += 1;
    formula_ = `=(${rangeexp.getAddress()}-${rangeexp_neg.getAddress()})/(${rangeexp.getAddress()}+${rangeexp_neg.getAddress()})`;
    formula_ = ones ? `=VSTACK(${formula_.substring(1)}, SEQUENCE(1,${number_of_cols})*0+1)` : formula_;
    let n_rows_withones = ones ? 1 : 0;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range = doConditionalFormatting(sheet, cursorrow, cursorcol, number_of_rows + 1 * n_rows_withones, number_of_cols, "CornflowerBlue", "Tomato", true, false, ones);
    cursorrow += number_of_rows + 1 * n_rows_withones;
    return { row: cursorrow, col: cursorcol, range };
}
function GELU(sheet: ExcelScript.Worksheet, header: string, cursorrow: number, cursorcol: number, in_range: ExcelScript.Range, number_of_rows: number, number_of_cols: number, ones: boolean = true, input_var_name?: string) {
    headerFormat(sheet, header, cursorrow, cursorcol);
    ;
    cursorrow++;
    const _eq = `✏️ u = √(2/𝜋)(${input_var_name}+.0447 ${input_var_name}³)`;
    writeEquation(sheet, _eq, cursorrow, cursorcol, "left");
    cursorrow += 1;
    let formula_ = ``;
    formula_ = `=0.797885*(${in_range.getAddress()}+0.044715*(${in_range.getAddress()})^3)`;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range1 = doConditionalFormatting(sheet, cursorrow, cursorcol, number_of_rows, number_of_cols, "ForestGreen", "PaleGreen");
    cursorrow += number_of_rows;
    const eq = `✏️ ${header.replace(/=.*/, "")} = (${input_var_name}/2)(1+Tanh(u))`;
    sheet.getCell(cursorrow, cursorcol).setValue(eq);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setSize(10);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setName("Bradley Hand");
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setBold(false);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setColor("red");
    sheet.getCell(cursorrow, cursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    sheet.getCell(cursorrow, cursorcol).getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.bottom);
    cursorrow++;
    formula_ = ``;
    formula_ =
        `=0.5*${in_range.getAddress()}*(1+TANH(${range1.getAddress()}))`;
    formula_ = ones ? `=VSTACK(${formula_.substring(1)}, SEQUENCE(1,${number_of_cols})*0+1)` : formula_;
    let n_rows_withones = ones ? 1 : 0;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range = doConditionalFormatting(sheet, cursorrow, cursorcol, number_of_rows + 1 * n_rows_withones, number_of_cols, "CornflowerBlue", "Tomato", true, false, ones);
    cursorrow += number_of_rows + 1 * n_rows_withones;
    return { row: cursorrow, col: cursorcol, range };
}
function GELUerf(sheet: ExcelScript.Worksheet, header: string, cursorrow: number, cursorcol: number, in_range: ExcelScript.Range, number_of_rows: number, number_of_cols: number, ones: boolean = true, input_var_name?: string, ind?: number) {
    headerFormat(sheet, header, cursorrow, cursorcol);
    cursorrow++;
    const _eq = `✏️ ${toSubscript("u_" + String(ind))} = 1+erf(${input_var_name}/√2)`;
    writeEquation(sheet, _eq, cursorrow, cursorcol, "left");
    cursorrow += 1;
    let formula_ = ``;
    formula_ = `=1+ERF(${in_range.getAddress()}/SQRT(2))`;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range1 = doConditionalFormatting(sheet, cursorrow, cursorcol, number_of_rows, number_of_cols, "ForestGreen", "PaleGreen");
    cursorrow += number_of_rows;
    const _eq2 = `✏️ ${header.replace(/=.*/, "")}= 0.5${input_var_name} ⊙ ${toSubscript("u" + String(ind))} `;
    writeEquation(sheet, _eq2, cursorrow, cursorcol, "left");
    cursorrow++;
    formula_ = ``;
    formula_ =
        `=0.5*${in_range.getAddress()}*${range1.getAddress()}`;
    formula_ = ones ? `=VSTACK(${formula_.substring(1)}, SEQUENCE(1,${number_of_cols})*0+1)` : formula_;
    let n_rows_withones = ones ? 1 : 0;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range = doConditionalFormatting(sheet, cursorrow, cursorcol, number_of_rows + 1 * n_rows_withones, number_of_cols, "CornflowerBlue", "Tomato", true, false, ones);
    cursorrow += number_of_rows + 1 * n_rows_withones;
    return { row: cursorrow, col: cursorcol, range };
}
function RandInput(workbook: ExcelScript.Workbook, sheet: ExcelScript.Worksheet, header: string, cursorrow: number, cursorcol: number, number_of_rows: number, number_of_cols: number, n_class: number, ones: boolean = true) {
    sheet.getCell(cursorrow, cursorcol).setValue(header);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setSize(12);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setBold(true);
    sheet.getCell(cursorrow, cursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    cursorrow += 1;
    let formula_ = ``;
    formula_ = `=RANDARRAY(${number_of_rows}, ${number_of_cols}, -1,1)`;
    formula_ = ones ? `=VSTACK(${formula_.substring(1)}, SEQUENCE(1,${number_of_cols})*0+1)` : formula_;
    let n_rows_withones = ones ? 1 : 0;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range = doConditionalFormatting(sheet, cursorrow, cursorcol, number_of_rows + 1 * n_rows_withones, number_of_cols, "CornflowerBlue", "Tomato", true, false, ones);
    let tempcursorrow = cursorrow;
    let tempcursorcol = cursorcol;
    tempcursorrow = tempcursorrow + number_of_rows + 1 * n_rows_withones;
    tempcursorrow++;
    sheet.getCell(tempcursorrow, tempcursorcol).setValue(`target = OHE(${n_class})`);
    sheet.getCell(tempcursorrow, tempcursorcol).getFormat().getFont().setSize(12);
    sheet.getCell(tempcursorrow, tempcursorcol).getFormat().getFont().setBold(true);
    sheet.getCell(tempcursorrow, tempcursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    sheet.getCell(tempcursorrow + 1, tempcursorcol - 1).setValue("Class ⟶  ");
    sheet.getCell(tempcursorrow + 1, tempcursorcol - 1).getFormat().getFont().setBold(false);
    sheet.getCell(tempcursorrow + 1, tempcursorcol - 1).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
    tempcursorrow++;
    let n_batch = number_of_cols;
    for (let b = 0; b < n_batch; b++) {
        let rand_cls = Math.floor(Math.random() * n_class);
        dropDownSequence(sheet, tempcursorrow, tempcursorcol + b, n_class, rand_cls);
        sheet.getCell(tempcursorrow, tempcursorcol + b)
            .getFormat()
            .getFill()
            .setColor("#f1f4ed");
        const formula_ohe = `=TRANSPOSE(--(SEQUENCE(1, ${n_class}) = 1+${convertToA1(tempcursorrow, tempcursorcol + b)}))`;
        sheet
            .getCell(tempcursorrow + 1, tempcursorcol + b)
            .setFormula(formula_ohe);
    }
    const range_ohe = doConditionalFormatting(sheet, tempcursorrow + 1, tempcursorcol, n_class, n_batch, "WhiteSmoke", "Red", true, false, false);
    sheet.addNamedItem("ohe", range_ohe);
    cursorrow += number_of_rows + 1 * n_rows_withones;
    cursorrow += n_class + 3;
    return { row: cursorrow, col: cursorcol, range };
}
function RandInput2(workbook: ExcelScript.Workbook, sheet: ExcelScript.Worksheet, header: string, cursorrow: number, cursorcol: number, number_of_rows: number, number_of_cols: number, n_class: number, ones: boolean = true, randorseq: string = "random", hasLoss: boolean = false, seed: number = 1) {
    headerFormat(sheet, header, cursorrow, cursorcol, "left", "black");
    cursorrow += 1;
    let n_rows_withones = ones ? 1 : 0;
    if (randorseq === "seeded") {
        let state = seed >>> 0;
        const rand = () => {
            state = (1664525 * state + 1013904223) >>> 0;
            return state / 4294967296;
        };
        let data: number[][] = [];
        for (let i = 0; i < number_of_rows; i++) {
            let row: number[] = [];
            for (let j = 0; j < number_of_cols; j++) {
                row.push(rand() * 2 - 1);
            }
            data.push(row);
        }
        if (ones) {
            data.push(new Array(number_of_cols).fill(1));
        }
        sheet.getRangeByIndexes(cursorrow, cursorcol, number_of_rows + n_rows_withones, number_of_cols)
            .setValues(data);
    }
    else {
        let formula_ = `=RANDARRAY(${number_of_rows}, ${number_of_cols}, -1, 1)`;
        formula_ = ones
            ? `=VSTACK(${formula_.substring(1)}, SEQUENCE(1,${number_of_cols})*0+1)`
            : formula_;
        sheet
            .getCell(cursorrow, cursorcol)
            .setFormula(formula_);
    }
    const range = doConditionalFormatting(sheet, cursorrow, cursorcol, number_of_rows + 1 * n_rows_withones, number_of_cols, "CornflowerBlue", "Tomato", true, false, ones);
    if (hasLoss) {
        let tempcursorrow = cursorrow;
        let tempcursorcol = cursorcol;
        tempcursorrow = tempcursorrow + number_of_rows + 1 * n_rows_withones;
        tempcursorrow++;
        const header = `target = OHE(class=${n_class})`;
        headerFormat(sheet, header, tempcursorrow, tempcursorcol, "left", "black");
        sheet.getCell(tempcursorrow + 1, tempcursorcol - 1).setValue("Class ⟶  ");
        sheet.getCell(tempcursorrow + 1, tempcursorcol - 1).getFormat().getFont().setBold(false);
        sheet.getCell(tempcursorrow + 1, tempcursorcol - 1).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
        tempcursorrow++;
        let n_batch = number_of_cols;
        for (let b = 0; b < n_batch; b++) {
            let rand_cls = Math.floor(Math.random() * n_class);
            dropDownSequence(sheet, tempcursorrow, tempcursorcol + b, n_class, rand_cls);
            sheet.getCell(tempcursorrow, tempcursorcol + b)
                .getFormat()
                .getFill()
                .setColor("#f1f4ed");
            const formula_ohe = `=TRANSPOSE(--(SEQUENCE(1, ${n_class}) = 1+${convertToA1(tempcursorrow, tempcursorcol + b)}))`;
            sheet
                .getCell(tempcursorrow + 1, tempcursorcol + b)
                .setFormula(formula_ohe);
        }
        const range_ohe = doConditionalFormatting(sheet, tempcursorrow + 1, tempcursorcol, n_class, n_batch, "WhiteSmoke", "Red", true, false, false);
        sheet.addNamedItem("ohe", range_ohe);
        cursorrow += n_class + 3;
    }
    cursorrow += number_of_rows + 1 * n_rows_withones;
    return { row: cursorrow, col: cursorcol, range };
}
function CELoss(sheet: ExcelScript.Worksheet, header: string, cursorrow: number, cursorcol: number, in_range: ExcelScript.Range, ohe_range: ExcelScript.Range, number_of_rows: number, number_of_cols: number, ones: boolean = true, input_var_name?: string) {
    headerFormat(sheet, header, cursorrow, cursorcol);
    ;
    cursorrow++;
    let formula_ = ``;
    formula_ = `=LN(${in_range.getAddress()}+1E-9)`;
    const _eq = `✏️ 𝓁 = ln(${input_var_name})`;
    writeEquation(sheet, _eq, cursorrow, cursorcol, "left");
    cursorrow++;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range_ln = doConditionalFormatting(sheet, cursorrow, cursorcol, number_of_rows, number_of_cols, "CornflowerBlue", "Tomato", true, false, ones);
    cursorrow += number_of_rows;
    formula_ = `=${ohe_range.getAddress()}*${range_ln.getAddress()}`;
    const _eq2 = `✏️ p = target. 𝓁`;
    writeEquation(sheet, _eq2, cursorrow, cursorcol, "left");
    cursorrow++;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range_yln = doConditionalFormatting(sheet, cursorrow, cursorcol, number_of_rows, number_of_cols, "CornflowerBlue", "Tomato", true, false, ones);
    cursorrow += number_of_rows;
    formula_ = `=-BYCOL(${range_yln.getAddress()}, LAMBDA(col, SUM(col)))`;
    const _eq3 = `✏️ ${header.replace(/=.*/, "")} = −∑p`;
    writeEquation(sheet, _eq3, cursorrow, cursorcol, "left");
    cursorrow++;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range = doConditionalFormatting(sheet, cursorrow, cursorcol, 1, number_of_cols, "CornflowerBlue", "Tomato", true, false, ones);
    cursorrow++;
    formula_ = `=AVERAGE(${range.getAddress()})`;
    const header2 = "mean(" + header.replace(/=.*/, "") + ")";
    headerFormat(sheet, header2, cursorrow, cursorcol);
    cursorrow++;
    const meanrange = sheet.getRangeByIndexes(cursorrow, cursorcol, 1, 2);
    meanrange.merge();
    meanrange.setFormula(formula_);
    meanrange.getFormat().getFont().setColor("yellow");
    meanrange.getFormat().getFont().setBold(true);
    meanrange.getFormat().getFont().setSize(9);
    meanrange.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    meanrange.getFormat().getFont().setName("Ink Free");
    doConditionalFormatting(sheet, cursorrow, cursorcol, 1, 2, "red", "red", true, false, ones);
    sheet.addNamedItem("LossMean", meanrange);
    cursorrow++;
    return { row: cursorrow, col: cursorcol, range };
}
function MSELoss(sheet: ExcelScript.Worksheet, header: string, cursorrow: number, cursorcol: number, in_range: ExcelScript.Range, ohe_range: ExcelScript.Range, number_of_rows: number, number_of_cols: number, ones: boolean = true, input_var_name?: string) {
    headerFormat(sheet, header, cursorrow, cursorcol);
    ;
    cursorrow++;
    let formula_ = ``;
    formula_ = `=(${in_range.getAddress()}-${ohe_range.getAddress()})`;
    const _eq = `✏️ d = (${input_var_name}−target)`;
    writeEquation(sheet, _eq, cursorrow, cursorcol, "left");
    cursorrow++;
    sheet.getCell(cursorrow, cursorcol).setFormula(formula_);
    const range_err = doConditionalFormatting(sheet, cursorrow, cursorcol, number_of_rows, number_of_cols, "CornflowerBlue", "Tomato", true, false, ones);
    cursorrow += number_of_rows;
    formula_ = `=${range_err.getAddress()}^2`;
    const _eq1 = `✏️ d²`;
    writeEquation(sheet, _eq1, cursorrow, cursorcol, "left");
    cursorrow++;
    sheet.getCell(cursorrow, cursorcol).setFormula(formula_);
    const range_sq = doConditionalFormatting(sheet, cursorrow, cursorcol, number_of_rows, number_of_cols, "CornflowerBlue", "Tomato", true, false, ones);
    cursorrow += number_of_rows;
    formula_ = `=BYCOL(${range_sq.getAddress()}, LAMBDA(col, AVERAGE(col)))`;
    const _eq2 = `✏️ ${header.replace(/=.*/, "")} = mean(d²)`;
    writeEquation(sheet, _eq2, cursorrow, cursorcol, "left");
    cursorrow++;
    sheet.getCell(cursorrow, cursorcol).setFormula(formula_);
    const range_colmean = doConditionalFormatting(sheet, cursorrow, cursorcol, 1, number_of_cols, "CornflowerBlue", "Tomato", true, false, ones);
    cursorrow++;
    formula_ = `=AVERAGE(${range_colmean.getAddress()})`;
    const header2 = "mean(" + header.replace(/=.*/, "") + ")";
    headerFormat(sheet, header2, cursorrow, cursorcol);
    cursorrow++;
    const meanrange = sheet.getRangeByIndexes(cursorrow, cursorcol, 1, 2);
    meanrange.merge();
    meanrange.setFormula(formula_);
    meanrange.getFormat().getFont().setColor("yellow");
    meanrange.getFormat().getFont().setBold(true);
    meanrange.getFormat().getFont().setSize(9);
    meanrange.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    meanrange.getFormat().getFont().setName("Ink Free");
    doConditionalFormatting(sheet, cursorrow, cursorcol, 1, 2, "red", "red", true, false, ones);
    sheet.addNamedItem("LossMean", meanrange);
    cursorrow++;
    return { row: cursorrow, col: cursorcol, range: range_colmean };
}
function LayerNorm(sheet: ExcelScript.Worksheet, header: string, cursorrow: number, cursorcol: number, in_range: ExcelScript.Range, number_of_rows: number, number_of_cols: number, ones: boolean = true, input_var_name?: string, ind?: number) {
    headerFormat(sheet, header, cursorrow, cursorcol);
    ;
    cursorrow += 1;
    let eqh = `✏️ ${toSubscript("µ_" + String(ind))} = mean(${input_var_name})`;
    writeEquation(sheet, eqh, cursorrow, cursorcol, "left");
    cursorrow += 1;
    let formula_ = ``;
    formula_ = `=BYCOL(${in_range.getAddress()}, LAMBDA(col, AVERAGE(col)))`;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range1 = doConditionalFormatting(sheet, cursorrow, cursorcol, 1, number_of_cols, "ForestGreen", "PaleGreen");
    cursorrow += 1;
    eqh = `️✏️ ${toSubscript("𝜎²_" + String(ind))}=var(${input_var_name})`;
    writeEquation(sheet, eqh, cursorrow, cursorcol, "left");
    cursorrow += 1;
    formula_ = `=BYCOL(${in_range.getAddress()}, LAMBDA(col, VAR.P(col)))`;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range2 = doConditionalFormatting(sheet, cursorrow, cursorcol, 1, number_of_cols, "ForestGreen", "PaleGreen");
    cursorrow += 1;
    eqh = `✏️ ${header.replace(/=.*/, "")} = (${input_var_name}－${toSubscript("µ_" + String(ind))})∕√(${toSubscript("𝜎²_" + String(ind))}+𝜀)`;
    writeEquation(sheet, eqh, cursorrow, cursorcol, "left");
    cursorrow += 1;
    formula_ = `=(${in_range.getAddress()} - ${range1.getAddress()})/SQRT(${range2.getAddress()} + 0.00000001)`;
    formula_ = ones ? `=VSTACK(${formula_.substring(1)}, SEQUENCE(1,${number_of_cols})*0+1)` : formula_;
    let n_rows_withones = ones ? 1 : 0;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range = doConditionalFormatting(sheet, cursorrow, cursorcol, number_of_rows + 1 * n_rows_withones, number_of_cols, "CornflowerBlue", "Tomato", true, false, ones);
    cursorrow += number_of_rows + 1 * n_rows_withones;
    return { row: cursorrow, col: cursorcol, range };
}
function L1Normalization(sheet: ExcelScript.Worksheet, header: string, cursorrow: number, cursorcol: number, in_range: ExcelScript.Range, number_of_rows: number, number_of_cols: number) {
    sheet.getCell(cursorrow, cursorcol).setValue("L1");
    sheet.getCell(cursorrow, cursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    cursorrow += 1;
    let formula_ = ``;
    formula_ = `=BYCOL(${in_range.getAddress()}, LAMBDA(col, SUMPRODUCT(ABS(col))))`;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range_ = doFormatting(sheet, cursorrow, cursorcol, 1, number_of_cols, "#fff2cc");
    cursorrow += 2;
    sheet.getCell(cursorrow, cursorcol).setValue(header);
    sheet.getCell(cursorrow, cursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    cursorrow += 1;
    formula_ = `=${in_range.getAddress()}/IF(${range_.getAddress()}=0,1, ${range_.getAddress()})`;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range = doFormatting(sheet, cursorrow, cursorcol, number_of_rows, number_of_cols, "#fff2cc");
    cursorrow += number_of_rows + 1;
    return { row: cursorrow, col: cursorcol, range };
}
function L2Normalization(sheet: ExcelScript.Worksheet, header: string, cursorrow: number, cursorcol: number, in_range: ExcelScript.Range, number_of_rows: number, number_of_cols: number) {
    sheet.getCell(cursorrow, cursorcol).setValue("L2");
    sheet.getCell(cursorrow, cursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    cursorrow += 1;
    let formula_ = ``;
    formula_ = `=BYCOL(${in_range.getAddress()}, LAMBDA(col, SQRT(SUM(col^2))))`;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range_ = doFormatting(sheet, cursorrow, cursorcol, 1, number_of_cols, "#ffa7cc");
    cursorrow += 2;
    sheet.getCell(cursorrow, cursorcol).setValue(header);
    sheet.getCell(cursorrow, cursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    cursorrow += 1;
    formula_ = `=${in_range.getAddress()}/IF(${range_.getAddress()}=0,1, ${range_.getAddress()})`;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range = doFormatting(sheet, cursorrow, cursorcol, number_of_rows, number_of_cols, "#ffa7cc");
    cursorrow += number_of_rows + 1;
    return { row: cursorrow, col: cursorcol, range };
}
function softmax(sheet: ExcelScript.Worksheet, header: string, cursorrow: number, cursorcol: number, in_range: ExcelScript.Range, number_of_rows: number, number_of_cols: number, ones: boolean = true, input_var_name?: string) {
    headerFormat(sheet, header, cursorrow, cursorcol);
    ;
    cursorrow += 1;
    const _eq = `✏️ 𝑒 = exp(${input_var_name})`;
    writeEquation(sheet, _eq, cursorrow, cursorcol, "left");
    cursorrow += 1;
    let formula_ = ``;
    formula_ = `=EXP(${in_range.getAddress()})`;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range1 = doConditionalFormatting(sheet, cursorrow, cursorcol, number_of_rows, number_of_cols, "ForestGreen", "PaleGreen");
    cursorrow += number_of_rows;
    const _eq2 = `✏️ ∑𝑒`;
    writeEquation(sheet, _eq2, cursorrow, cursorcol, "left");
    cursorrow += 1;
    formula_ = `=BYCOL(${range1.getAddress()}, LAMBDA(col, SUM(col)))`;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range2 = doConditionalFormatting(sheet, cursorrow, cursorcol, 1, number_of_cols, "ForestGreen", "PaleGreen");
    cursorrow += 1;
    const _eq3 = `✏️ ${header.replace(/=.*/, "")} = 𝑒∕∑𝑒`;
    writeEquation(sheet, _eq3, cursorrow, cursorcol, "left");
    cursorrow += 1;
    formula_ = `=${range1.getAddress()}/${range2.getAddress()}`;
    formula_ = ones ? `=VSTACK(${formula_.substring(1)}, SEQUENCE(1,${number_of_cols})*0+1)` : formula_;
    let n_rows_withones = ones ? 1 : 0;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range = doConditionalFormatting(sheet, cursorrow, cursorcol, number_of_rows + 1 * n_rows_withones, number_of_cols, "CornflowerBlue", "Tomato", true, false, ones);
    cursorrow += number_of_rows + 1 * n_rows_withones;
    return { row: cursorrow, col: cursorcol, range };
}
function drawXwithText(sheet: ExcelScript.Worksheet, header: string, extractedString: string, cursorrow: number, cursorcol: number, number_of_rows: number, number_of_cols: number) {
    sheet.getCell(cursorrow, cursorcol).setValue(header);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setSize(12);
    sheet.getCell(cursorrow, cursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    cursorrow += 1;
    const charCount = extractedString.length;
    for (let i = 0; i < charCount; i++) {
        const cell = sheet.getCell(cursorrow, cursorcol + i);
        cell.setValue(extractedString[i]);
    }
    const range2 = doFormatting(sheet, cursorrow, cursorcol, number_of_rows, number_of_cols, "#f6b26b");
    cursorrow += number_of_rows + 1;
    sheet.getCell(cursorrow, cursorcol).setValue("Token ID");
    sheet.getCell(cursorrow, cursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    cursorrow += 1;
    const formula_ = `=CODE(LOWER(${range2.getAddress()}))- CODE("a")`;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range = doFormatting(sheet, cursorrow, cursorcol, number_of_rows, number_of_cols, "white");
    cursorrow += number_of_rows + 1;
    return { row: cursorrow, col: cursorcol, range };
}
function transposeMatrix(sheet: ExcelScript.Worksheet, in_range: ExcelScript.Range, header: string, var_name: string, cursorrow: number, cursorcol: number, number_of_rows: number, number_of_cols: number) {
    sheet.getCell(cursorrow, cursorcol).setValue(header);
    sheet.getCell(cursorrow, cursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    cursorrow += 1;
    const formula_ = `=TRANSPOSE(${in_range.getAddress()})`;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range = doFormatting(sheet, cursorrow, cursorcol, number_of_rows, number_of_cols, "#daf2d1");
    cursorrow += number_of_rows + 1;
    return { row: cursorrow, col: cursorcol, range };
}
function add2Matrices(sheet: ExcelScript.Worksheet, range1: ExcelScript.Range, range2: ExcelScript.Range, header: string, cursorrow: number, cursorcol: number, number_of_rows: number, number_of_cols: number) {
    sheet.getCell(cursorrow, cursorcol).setValue(header);
    sheet.getCell(cursorrow, cursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    cursorrow += 1;
    const formula_ = `=${range1.getAddress()}+${range2.getAddress()}`;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range = doFormatting(sheet, cursorrow, cursorcol, number_of_rows, number_of_cols, "#cfe2f3");
    cursorrow += number_of_rows + 1;
    return { row: cursorrow, col: cursorcol, range };
}
function multMatricesWithRand(sheet: ExcelScript.Worksheet, range1: ExcelScript.Range, range2: ExcelScript.Range, indim: number, outdim: number, outcols: number, header: string, cursorrow: number, cursorcol: number, ones: boolean = false) {
    const number_of_rows = outdim;
    const number_of_cols = indim;
    sheet.getCell(cursorrow, cursorcol).setValue(header);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setSize(12);
    sheet.getCell(cursorrow, cursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    cursorrow += 1;
    const formula_ = `=MMULT(${range1.getAddress()},${range2.getAddress()})`;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range = doConditionalFormatting(sheet, cursorrow, cursorcol, number_of_rows, outcols, "CornflowerBlue", "Tomato", true, false, ones);
    cursorrow += number_of_rows + 1;
    return { row: cursorrow, col: cursorcol, range };
}
function multMatricesWithRand2(sheet: ExcelScript.Worksheet, range1: ExcelScript.Range, range2: ExcelScript.Range, indim: number, outdim: number, outcols: number, header: string, cursorrow: number, cursorcol: number, ones: boolean = false, linear_eq?: string, input_var_name?: string) {
    const number_of_rows = outdim;
    const number_of_cols = indim;
    cursorrow -= 1;
    headerFormat(sheet, header, cursorrow, cursorcol);
    ;
    cursorrow++;
    const _eq = `✏️ ${header.replace(/=.*/, "")} = ${linear_eq}`;
    writeEquation(sheet, _eq, cursorrow, cursorcol, "left");
    cursorrow++;
    let formula_ = ``;
    formula_ = `=MMULT(${range1.getAddress()},${range2.getAddress()})`;
    formula_ = ones ? `=VSTACK(${formula_.substring(1)}, SEQUENCE(1,${outcols})*0+1)` : formula_;
    let n_rows_withones = ones ? 1 : 0;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range = doConditionalFormatting(sheet, cursorrow, cursorcol, number_of_rows + 1 * n_rows_withones, outcols, "CornflowerBlue", "Tomato", true, false, ones);
    cursorrow += number_of_rows + 1 * n_rows_withones;
    return { row: cursorrow, col: cursorcol, range };
}
function getDatasetInfo(workbook: ExcelScript.Workbook, datasetName: string) {
    const info_range = workbook.getNamedItem(datasetName + "Info").getRange();
    const values = info_range.getValues();
    let dict: {
        [key: string]: string | number;
    } = {};
    for (let i = 0; i < values.length; i++) {
        const key = String(values[i][0]);
        const val = values[i][1];
        dict[key] = (typeof val === "number") ? val : String(val);
    }
    return dict;
}
function multTwoMatrices2(sheet: ExcelScript.Worksheet, header: string, cursorrow: number, cursorcol: number, matrix1: string, transpose1: boolean = false, matrix2: string, transpose2: boolean = false, showMatrix2: boolean = true, matrix1row: number, matrix1col: number, matrix2row: number, matrix2col: number, ones: boolean = false) {
    let outmatrixrow = matrix1row;
    let outmatrixcol = matrix1col;
    let formula_ = ``;
    let formula_m1 = ``;
    let formula_m2 = ``;
    let matrix1header = `${toSubscript(matrix1)}`;
    let matrix2header = `${toSubscript(matrix2)}`;
    formula_ = `=MMULT(${matrix1},${matrix2})`;
    let tempcursorrow = 0;
    let tempcursorcol = 0;
    if (showMatrix2) {
        formula_m2 = `=${matrix2}`;
        sheet.getCell(cursorrow, cursorcol).setValue(matrix2header);
        sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setSize(12);
        sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setBold(true);
        sheet.getCell(cursorrow, cursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
        cursorrow++;
        tempcursorrow = cursorrow;
        tempcursorcol = cursorcol;
        sheet
            .getCell(cursorrow, cursorcol)
            .setFormula(formula_m2);
        const range_m2 = doConditionalFormatting(sheet, cursorrow, cursorcol, matrix2row, matrix2col, "CornflowerBlue", "Tomato", true, false, false);
        cursorrow += matrix2row + 1;
    }
    if (!transpose1 && !transpose2) {
        formula_m1 = `=${matrix1}`;
        cursorcol -= (matrix1col + 1);
        matrix1header = matrix1header;
        sheet.getCell(cursorrow, cursorcol).setValue(matrix1header);
        sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setSize(12);
        sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setBold(true);
        sheet.getCell(cursorrow, cursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
        cursorrow++;
        if (showMatrix2) {
            formula_ = `=MMULT(${convertToA1(cursorrow, cursorcol)}#, ${convertToA1(tempcursorrow, tempcursorcol)}#)`;
        }
        else {
            formula_ = `=MMULT(${convertToA1(cursorrow, cursorcol)}#, ${matrix2})`;
        }
        sheet
            .getCell(cursorrow, cursorcol)
            .setFormula(formula_m1);
        sheet
            .getCell(cursorrow, cursorcol)
            .setFormula(formula_m1);
        const range_m1 = doConditionalFormatting(sheet, cursorrow, cursorcol, matrix1row, matrix1col, "CornflowerBlue", "Tomato", false, true, false);
        cursorcol += matrix1col + 1;
    }
    else if (transpose1) {
        outmatrixrow = outmatrixcol;
        formula_m1 = `=TRANSPOSE(${matrix1})`;
        cursorcol -= (matrix1row + 1);
        matrix1header = matrix1header + "ᵀ";
        sheet.getCell(cursorrow, cursorcol).setValue(matrix1header);
        sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setSize(12);
        sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setBold(true);
        sheet.getCell(cursorrow, cursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
        cursorrow++;
        if (showMatrix2) {
            formula_ = `=MMULT(${convertToA1(cursorrow, cursorcol)}#, ${convertToA1(tempcursorrow, tempcursorcol)}#)`;
        }
        else {
            formula_ = `=MMULT(${convertToA1(cursorrow, cursorcol)}#, ${matrix2})`;
        }
        sheet
            .getCell(cursorrow, cursorcol)
            .setFormula(formula_m1);
        const range_m1 = doConditionalFormatting(sheet, cursorrow, cursorcol, matrix1col, matrix1row, "CornflowerBlue", "Tomato", false, true, false);
        cursorcol += matrix1row + 1;
    }
    else if (transpose2) {
        formula_ = `=MMULT(${matrix1},TRANSPOSE(${matrix2}))`;
        outmatrixcol = outmatrixrow;
    }
    cursorrow--;
    sheet.getCell(cursorrow, cursorcol).setValue(header + matrix1header + "." + matrix2header);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setSize(12);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setBold(true);
    sheet.getCell(cursorrow, cursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    cursorrow++;
    formula_ = ones ? `=VSTACK(${formula_.substring(1)}, SEQUENCE(1,${outmatrixcol})*0+1)` : formula_;
    let n_rows_withones = ones ? 1 : 0;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range = doConditionalFormatting(sheet, cursorrow, cursorcol, outmatrixrow + 1 * n_rows_withones, outmatrixcol, "CornflowerBlue", "Tomato", true, false, ones);
    cursorrow += transpose1 ? outmatrixcol + 1 * n_rows_withones : outmatrixrow + 1 * n_rows_withones;
    return { row: cursorrow, col: cursorcol, range };
}
function mult2Matrices(sheet: ExcelScript.Worksheet, range1: ExcelScript.Range, range2: ExcelScript.Range, indim: number, outdim: number, outcols: number, header: string, cursorrow: number, cursorcol: number) {
    const number_of_rows = outdim;
    const number_of_cols = indim;
    sheet.getCell(cursorrow, cursorcol).setValue(header);
    sheet.getCell(cursorrow, cursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    cursorrow += 1;
    const formula_ = `=MMULT(${range1.getAddress()},${range2.getAddress()})`;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range = doConditionalFormatting(sheet, cursorrow, cursorcol, number_of_rows, outcols);
    cursorrow += number_of_rows + 1;
    return { row: cursorrow, col: cursorcol, range };
}
function scaleMatrix(sheet: ExcelScript.Worksheet, header: string, cursorrow: number, cursorcol: number, scaleFactor: number, multordiv: string = "mult", matrix: string, number_of_rows: number, number_of_cols: number, ones: boolean = false) {
    sheet.getCell(cursorrow, cursorcol).setValue(header);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setSize(12);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setBold(true);
    sheet.getCell(cursorrow, cursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    cursorrow += 1;
    let formula_ = ``;
    formula_ = `=${matrix}*${scaleFactor}`;
    if (multordiv == "div") {
        formula_ = `=${matrix}*1/${scaleFactor}`;
    }
    formula_ = ones ? `=VSTACK(${formula_.substring(1)}, SEQUENCE(1,${number_of_cols})*0+1)` : formula_;
    let n_rows_withones = ones ? 1 : 0;
    sheet
        .getCell(cursorrow, cursorcol)
        .setFormula(formula_);
    const range = doConditionalFormatting(sheet, cursorrow, cursorcol, number_of_rows + 1 * n_rows_withones, number_of_cols, "CornflowerBlue", "Tomato", true, false, ones);
    cursorrow += number_of_rows + 1 * n_rows_withones;
    return { row: cursorrow, col: cursorcol, range };
}
function changeTitle(ops_string: string, var_string: string) {
    let out_string = "";
    out_string = ops_string + "(" + var_string + ")";
    return out_string;
}
function changeTitle2(out_var: string, ops_string: string, var_string: string) {
    let out_string = "";
    out_string = out_var + " = " + ops_string + "(" + var_string + ")";
    return out_string;
}
function drawShape(sheet: ExcelScript.Worksheet, text: string, posW: number, posH: number, shapeW: number, shapeH: number) {
    const shape = sheet.addGeometricShape(ExcelScript.GeometricShapeType.roundRectangle);
    shape.setName("in");
    shape.setLeft(posW);
    shape.setTop(posH);
    shape.setWidth(shapeW);
    shape.setHeight(shapeH);
    shape.getFill().setSolidColor("lightblue");
    shape.getTextFrame().getTextRange().setText(text);
    shape.getTextFrame().getTextRange().getFont().setColor("black");
    shape.getTextFrame().getTextRange().getFont().setSize(8);
    shape.getTextFrame().setHorizontalAlignment(ExcelScript.ShapeTextHorizontalAlignment.center);
    shape.getTextFrame().setVerticalAlignment(ExcelScript.ShapeTextVerticalAlignment.middle);
    shape.getLineFormat().setWeight(0);
}
function drawObject(sheet: ExcelScript.Worksheet, text: string, posW: number, posH: number, shapeW: number, shapeH: number, previousShape?: ExcelScript.Shape) {
    const shape = sheet.addGeometricShape(ExcelScript.GeometricShapeType.roundRectangle);
    shape.setLeft(posW);
    shape.setTop(posH);
    shape.setWidth(shapeW);
    shape.setHeight(shapeH);
    shape.getFill().setSolidColor("lightblue");
    shape.getTextFrame().getTextRange().setText(text);
    shape.getTextFrame().getTextRange().getFont().setColor("black");
    shape.getTextFrame().getTextRange().getFont().setSize(8);
    shape.getTextFrame().setHorizontalAlignment(ExcelScript.ShapeTextHorizontalAlignment.center);
    shape.getTextFrame().setVerticalAlignment(ExcelScript.ShapeTextVerticalAlignment.middle);
    return shape;
}
function shapeConnector(sheet: ExcelScript.Worksheet, shape0name: string, shape1name: string) {
    const allShapes = sheet.getShapes();
    const previousShape = allShapes.find(s => s.getName() === shape0name);
    const shape = allShapes.find(s => s.getName() === shape1name);
    const arrow1 = sheet.addLine(0, 0, 10, 10, ExcelScript.ConnectorType.curve);
    arrow1.getLineFormat().setColor("black");
    arrow1.getLineFormat().setWeight(1);
    arrow1.getLineFormat().setDashStyle(ExcelScript.ShapeLineDashStyle.systemDot);
    arrow1.getLine().setBeginArrowheadStyle(ExcelScript.ArrowheadStyle.oval);
    arrow1.getLine().connectBeginShape(previousShape, 3);
    arrow1.getLine().setEndArrowheadStyle(ExcelScript.ArrowheadStyle.stealth);
    arrow1.getLine().connectEndShape(shape, 1);
}
function drawBlockBorder(blockID: number, sheet: ExcelScript.Worksheet, input_var_name: string, blockheader: string, cursorrow: number, cursorcol: number) {
    const number_of_cols = 20;
    const number_of_rows = 1;
    sheet.getCell(cursorrow, cursorcol).setValue("# " + (blockID + 1));
    sheet.getCell(cursorrow, cursorcol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    sheet.getCell(cursorrow, cursorcol + 1).setValue(" " + blockheader);
    sheet.getCell(cursorrow, cursorcol + 1).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    sheet.getCell(cursorrow, cursorcol + 1).getFormat().getFont().setColor("gray");
    sheet.getCell(cursorrow, cursorcol + 1).getFormat().getFont().setItalic(true);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setSize(8);
    sheet.getCell(cursorrow, cursorcol).getFormat().getFont().setColor("white");
    sheet.getCell(cursorrow, cursorcol).getFormat().getFill().setColor("black");
    const range = sheet.getRange(convertToA1(cursorrow, cursorcol) + ":" + convertToA1(cursorrow + number_of_rows - 1, cursorcol + number_of_cols - 1));
    const fmt = range.getFormat();
    const borders = [
        ExcelScript.BorderIndex.edgeTop,
    ];
    for (let b of borders) {
        fmt.getRangeBorder(b).setStyle(ExcelScript.BorderLineStyle.none);
    }
    return range;
}
function doneMessage(blockID: number, sheet: ExcelScript.Worksheet, colID: number, backprop: boolean = false, hlink: string = "") {
    colID = colID + 1;
    const rowID = blockID + 1;
    if (!backprop) {
        if (hlink == "") {
            sheet.getCell(rowID, colID).setValue("↓");
            sheet.getCell(rowID, colID).getFormat().getFont().setSize(10);
            sheet.getCell(rowID, colID).getFormat().getFont().setColor("green");
        }
        else {
            sheet.getCell(rowID, colID).setFormula(`=HYPERLINK("#Blueprint!${hlink}", "↓")`);
            sheet.getCell(rowID, colID).getFormat().getFont().setSize(10);
            sheet.getCell(rowID, colID).getFormat().getFont().setColor("green");
        }
    }
    else {
        sheet.getCell(rowID, colID + 1).setValue("↑");
        sheet.getCell(rowID, colID + 1).getFormat().getFont().setSize(10);
        sheet.getCell(rowID, colID + 1).getFormat().getFont().setColor("red");
    }
}
function fullCompleteMessage(sheet: ExcelScript.Worksheet, rowID: number, colID: number, mergeCells: boolean = false) {
    let msgCell: ExcelScript.Range;
    if (mergeCells) {
        msgCell = sheet.getRangeByIndexes(rowID, colID, 1, 2);
        msgCell.merge();
    }
    else {
        msgCell = sheet.getRangeByIndexes(rowID, colID, 1, 1);
    }
    msgCell.setFormula(`=HYPERLINK("#Blueprint!X", "✓ Blueprint Tab")`);
    msgCell.getFormat().getFont().setColor("Maroon");
    msgCell.getFormat().getFont().setName("Ink Free");
    msgCell.getFormat().getFont().setSize(8);
    msgCell.getFormat().getFont().setBold(true);
}
function addDashboardAnchor(workbook: ExcelScript.Workbook, sheet: ExcelScript.Worksheet, name: string, rowID: number, colID: number) {
    const anchorRange = sheet.getRangeByIndexes(rowID, colID, 1, 1);
    updateName(workbook, { row: rowID, col: colID, range: anchorRange }, name, sheet);
}
function drawDashboardBlock(sheet: ExcelScript.Worksheet, rowID: number, colID: number, height: number, width: number, title: string, listReference: string = "", titleColor: string = "Red") {
    const block = sheet.getRangeByIndexes(rowID, colID, height, width);
    block.getFormat().getFill().setColor("White");
    const fmt = block.getFormat();
    const borders = [
        ExcelScript.BorderIndex.edgeTop,
        ExcelScript.BorderIndex.edgeBottom,
        ExcelScript.BorderIndex.edgeLeft,
        ExcelScript.BorderIndex.edgeRight,
    ];
    for (const b of borders) {
        fmt.getRangeBorder(b).setStyle(ExcelScript.BorderLineStyle.continuous);
        fmt.getRangeBorder(b).setWeight(1);
        fmt.getRangeBorder(b).setColor("#A6A6A6");
    }
    const headerFill = "#FFC000";
    const headerBand = sheet.getRangeByIndexes(rowID, colID, 1, width);
    headerBand.getFormat().getFill().setColor(headerFill);
    headerBand.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    headerBand.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(ExcelScript.BorderLineStyle.continuous);
    headerBand.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setWeight(1);
    headerBand.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setColor("#D6A100");
    const hasBackLink = listReference !== "";
    const headerWidth = Math.max(1, width - (hasBackLink ? 4 : 2));
    const header = sheet.getRangeByIndexes(rowID, colID + 1, 1, headerWidth);
    if (headerWidth > 1) {
        header.merge();
    }
    header.setValue(title);
    header.getFormat().getFont().setName("Ink Free");
    header.getFormat().getFont().setSize(15);
    header.getFormat().getFont().setBold(true);
    header.getFormat().getFont().setColor(titleColor);
    header.getFormat().getFill().setColor(headerFill);
    header.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    header.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    if (hasBackLink) {
        linkedMessage(sheet, "↩ List", listReference, rowID, colID + width - 3, true, 8, 2, "SeaGreen");
        const backLinkCell = sheet.getRangeByIndexes(rowID, colID + width - 3, 1, 2);
        backLinkCell.getFormat().getFill().setColor(headerFill);
    }
}
function linkedMessage(sheet: ExcelScript.Worksheet, message: string, reference: string, rowID: number, colID: number, mergeCells: boolean = false, fnt: number = 8, mrgNumber: number = 2, clr: string = "Maroon", fillclr: string = "White") {
    let msgCell: ExcelScript.Range;
    if (mergeCells) {
        msgCell = sheet.getRangeByIndexes(rowID, colID, 1, mrgNumber);
        msgCell.merge();
    }
    else {
        msgCell = sheet.getRangeByIndexes(rowID, colID, 1, 1);
    }
    msgCell.setFormula(`=HYPERLINK("#${reference}", "${message}")`);
    msgCell.getFormat().getFont().setColor(clr);
    msgCell.getFormat().getFont().setName("Ink Free");
    msgCell.getFormat().getFont().setSize(fnt);
    msgCell.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    msgCell.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    msgCell.getFormat().getFont().setBold(false);
}
function lookUp() {
    const lookup: {
        [key: string]: [
            number,
            number,
            number
        ];
    } = {};
    lookup["A"] = [0.082, 0.078, 4.338];
    lookup["B"] = [0.015, 0.02, 2.849];
    lookup["C"] = [0.028, 0.04, 2.849];
    lookup["D"] = [0.043, 0.038, 2.964];
    lookup["E"] = [0.127, 0.11, 3.383];
    lookup["F"] = [0.022, 0.014, 2.612];
    lookup["G"] = [0.02, 0.03, 2.494];
    lookup["H"] = [0.061, 0.023, 2.612];
    lookup["I"] = [0.07, 0.086, 3.383];
    lookup["J"] = [0.0015, 0.0021, 2.964];
    lookup["K"] = [0.0077, 0.0097, 3.545];
    lookup["L"] = [0.04, 0.053, 4.338];
    lookup["M"] = [0.024, 0.027, 3.772];
    lookup["N"] = [0.067, 0.072, 3.196];
    lookup["O"] = [0.075, 0.061, 4.043];
    lookup["P"] = [0.019, 0.028, 4.889];
    lookup["Q"] = [0.00095, 0.0019, 4.889];
    lookup["R"] = [0.06, 0.073, 2.934];
    lookup["S"] = [0.063, 0.087, 3.545];
    lookup["T"] = [0.091, 0.067, 2.707];
    lookup["U"] = [0.028, 0.033, 2.934];
    lookup["V"] = [0.0098, 0.01, 2.732];
    lookup["W"] = [0.024, 0.0091, 4.043];
    lookup["X"] = [0.0015, 0.0027, 3.196];
    lookup["Y"] = [0.02, 0.016, 2.707];
    lookup["Z"] = [0.00074, 0.0044, 3.772];
    return lookup;
}
function dropDown(range: ExcelScript.Range, rowID: number, colID: number) {
    const options = ["TRUE", "FALSE"];
    const validation = range.getDataValidation();
    validation.setRule({
        list: {
            inCellDropDown: true,
            source: options.join(",")
        }
    });
    range.getCell(rowID, colID).setValue("TRUE");
}
function dropDownSequence(sheet: ExcelScript.Worksheet, rowID: number, colID: number, n: number, defaultValue: number = 0) {
    const cell = sheet.getCell(rowID, colID);
    const maxValue = n - 1;
    const options: string[] = [];
    for (let i = 0; i <= maxValue; i++) {
        options.push(i.toString());
    }
    const validation = cell.getDataValidation();
    validation.setRule({
        list: {
            inCellDropDown: true,
            source: options.join(",")
        }
    });
    cell.setValue(defaultValue);
}
function writeCode(sheet: ExcelScript.Worksheet, cursorrow: number, cursorcol: number, number_of_rows: number, number_of_cols: number) {
    const range = sheet.getRangeByIndexes(cursorrow, cursorcol, number_of_rows, number_of_cols);
    let values = range.getValues()[0];
    let numbersStr = values.map(v => v.toString()).join(", ");
    let resultStr = `X = torch.tensor([${numbersStr}])`;
    console.log(resultStr);
}
function writeTorchCode(sheet: ExcelScript.Worksheet, code: string, items: string[] = [], cursorrow: number, cursorcol: number, number_of_rows?: number, number_of_cols?: number, bgColor: string = "#2B2B2B", textColor: string = "#F0F0F0", itemLinks: string[] = [], itemAnchorNames: string[][] = []) {
    cursorcol += 4;
    const gutterWidth = 1;
    const markerWidth = 2;
    const separatorWidth = 1;
    const codeWidth = 22;
    function applyCodeFmt(fmt: ExcelScript.RangeFormat): void {
        fmt.getFill().setColor("#F6F8FC");
        fmt.getFont().setName("Courier New");
        fmt.getFont().setColor("#1F2937");
        fmt.getFont().setSize(13);
        fmt.getFont().setBold(false);
        fmt.setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
        fmt.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    }
    function applyMarkerFmt(fmt: ExcelScript.RangeFormat): void {
        fmt.getFill().setColor("#E5EBF5");
        fmt.getFont().setName("Courier New");
        fmt.getFont().setColor("#355C9A");
        fmt.getFont().setSize(11);
        fmt.getFont().setBold(false);
        fmt.setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
        fmt.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    }
    function applySeparatorFmt(fmt: ExcelScript.RangeFormat): void {
        fmt.getFill().setColor("#E5EBF5");
        fmt.getFont().setName("Courier New");
        fmt.getFont().setColor("#355C9A");
        fmt.getFont().setSize(11);
        fmt.getFont().setBold(false);
        fmt.setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
        fmt.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    }
    function applyGutterFmt(fmt: ExcelScript.RangeFormat): void {
        fmt.getFill().setColor("#EEF2F8");
    }
    function setCodeCell(cell: ExcelScript.Range, displayText: string): void {
        cell.merge(false);
        cell.setValue(displayText);
    }
    function setCodeAnchors(cell: ExcelScript.Range, anchorNames: string[] = []): void {
        for (const anchorName of anchorNames) {
            if (anchorName === "") {
                continue;
            }
            const codeAnchorName = getTorchCodeAnchorName(anchorName);
            const existingName = sheet.getNames().find(n => n.getName() === codeAnchorName);
            if (!existingName) {
                sheet.addNamedItem(codeAnchorName, cell);
            }
        }
    }
    function setMarkerCell(cell: ExcelScript.Range, linkName: string = ""): void {
        cell.merge(false);
        if (linkName !== "") {
            try {
                sheet.getNamedItem(linkName).getRange();
                cell.setFormula(`=HYPERLINK("#${sheet.getName()}!${linkName}","view")`);
                return;
            }
            catch {
            }
        }
        cell.setValue("");
    }
    function setSeparatorCell(cell: ExcelScript.Range): void {
        cell.merge(false);
        cell.setValue(">>");
    }
    function writeLine(row: number, displayText: string, linkName: string = "", anchorNames: string[] = []): void {
        const gutterCell = sheet.getRangeByIndexes(row, cursorcol - markerWidth - separatorWidth - gutterWidth, 1, gutterWidth);
        const markerCell = sheet.getRangeByIndexes(row, cursorcol - separatorWidth - markerWidth, 1, markerWidth);
        const separatorCell = sheet.getRangeByIndexes(row, cursorcol - separatorWidth, 1, separatorWidth);
        const codeCell = sheet.getRangeByIndexes(row, cursorcol, 1, codeWidth);
        setMarkerCell(markerCell, linkName);
        setSeparatorCell(separatorCell);
        setCodeCell(codeCell, displayText);
        setCodeAnchors(codeCell, anchorNames);
        applyMarkerFmt(markerCell.getFormat());
        applySeparatorFmt(separatorCell.getFormat());
        applyCodeFmt(codeCell.getFormat());
    }
    if (items.length > 0) {
        for (let _r = 0; _r < items.length; _r++) {
            writeLine(cursorrow + _r, items[_r], itemLinks[_r] || "", itemAnchorNames[_r] || []);
        }
    }
    if (code !== "") {
        writeLine(cursorrow, code, "", []);
    }
}
function getCode(range: ExcelScript.Range, input_var_name: string, output_var_name: string, cursorrow: number, cursorcol: number, number_of_rows: number, number_of_cols: number) {
    let values = range.getValues()[0];
    let numbersStr = values.map(v => v.toString()).join(", ");
    let resultStr = `${output_var_name} = torch.tensor([${numbersStr}])`;
    return resultStr;
}
function genTorchCode(func: string, assignVar: string = "", p: string = "", q: string = "") {
    let dictCode: {
        [key: string]: string;
    } = {};
    const _in = `torch.randn(${p},${q})`;
    const _ohe = `torch.randn(${p})`;
    const _ohe2 = `torch.randint(0,${p},(${q},))`;
    const _celoss = `loss_fn=nn.CrossEntropyLoss()`;
    const _mseloss = `loss_fn=nn.MSELoss()`;
    const _linear = `nn.Linear(${p},${q})`;
    const _relu = `nn.ReLU()`;
    const _tanh = `nn.Tanh()`;
    const _gelu = `nn.GELU()`;
    const _LN = `nn.LayerNorm(${p})`;
    const _soft = `nn.Softmax(dim=1)`;
    const _model = `model=nn.Sequential(${p})`;
    const _pred = `model(${p})`;
    const _loss = `loss_fn(${p},${q})`;
    const _in_image = `torch.randn(${p},1,${q},${q})`;
    const _ohe_image = `torch.randint(0,${p},(${q},))`;
    const _flatten = `nn.Flatten()`;
    dictCode["Rand"] = _in;
    dictCode["ohe"] = _ohe2;
    dictCode["CELoss"] = _celoss;
    dictCode["MSELoss"] = _mseloss;
    dictCode["ReLU"] = _relu;
    dictCode["Tanh"] = _tanh;
    dictCode["GELU"] = _gelu;
    dictCode["LayerNorm"] = _LN;
    dictCode["Softmax"] = _soft;
    dictCode["Linear"] = _linear;
    dictCode["Model"] = _model;
    dictCode["Predict"] = _pred;
    dictCode["loss"] = _loss;
    dictCode["Image"] = _in_image;
    dictCode["oheImage"] = _ohe_image;
    dictCode["Flatten"] = _flatten;
    let codeline: string = ``;
    if (assignVar == "")
        codeline = dictCode[func];
    else
        codeline = assignVar + "=" + dictCode[func];
    return codeline;
}
function genzForwardPassCode(fnc_body: string, fnc_x: string, fnc_return: string) {
    const code = `${fnc_return}=${fnc_body}(${fnc_x})`;
    return code;
}
function toSubscript(input: string): string {
    const subMap: {
        [key: string]: string;
    } = {
        "0": "₀",
        "1": "₁",
        "2": "₂",
        "3": "₃",
        "4": "₄",
        "5": "₅",
        "6": "₆",
        "7": "₇",
        "8": "₈",
        "9": "₉",
    };
    return input.replace(/_(\d+)/g, (_, digits) => digits
        .split("")
        .map(d => subMap[d] ?? d)
        .join(""));
}
