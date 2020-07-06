import * as xlsx from "node-xlsx";
import * as fs from "fs";
import * as path from "path";
interface I_in_out {
    "input": string,
    "output": string
}
console.log("\n");
let config: I_in_out[] = require("./config.json");

for (let one of config) {
    let inputfiles: string[] = [];
    try {
        inputfiles = fs.readdirSync(one.input);
    } catch (e) {
        console.error("directory", "非法输入路径", one.input);
        continue;
    }
    try {
        fs.readdirSync(one.output);
    } catch (e) {
        console.error("directory", "非法输出路径", one.output);
        continue;
    }
    inputfiles.forEach((filename) => {
        if (filename[0] === "~") {
            return;
        }
        if (!filename.endsWith(".xlsx")) {
            return;
        }
        let intputfilepath = path.join(one.input, filename);
        let buff = fs.readFileSync(intputfilepath);
        parseBuffToJson(buff, one.output, path.basename(filename, '.xlsx'));
        console.log("---->>>", one.input, " ", filename);
    });
}

console.log("\n");


function parseBuffToJson(buff: Buffer, outputDir: string, filename: string) {
    let sheets = xlsx.parse(buff);
    for (let one of sheets) {
        let lists: any = one.data;
        if (lists.length === 0) {
            continue;
        }
        let keyarr = lists[0];
        let typearr = lists[1];
        let obj: any = {};
        for (let i = 3; i < lists.length; i++) {
            if (lists[i][0] === undefined) {
                continue;
            }
            obj[lists[i][0]] = createObj(keyarr, typearr, lists[i]);
        }
        fs.writeFileSync(path.join(outputDir, filename + "_" + one.name + ".json"), JSON.stringify(obj, null, 4));
    }
}
function createObj(keyarr: string[], typearr: string[], dataarr: any[]) {
    let obj: any = {};
    for (let i = 0; i < keyarr.length; i++) {
        obj[keyarr[i]] = changeValue(dataarr[i], typearr[i]);
    }
    return obj;
}

function changeValue(value: any, type: string) {
    if (value === undefined) {
        value = "";
    }
    if (type === "bool") {
        if (typeof value === "string") {
            value = value.trim();
        }
        if (value === 0 || value == "" || value === "false") {
            return false;
        } else {
            return true;
        }
    } else if (type === "string") {
        return value.toString();
    } else if (type === "float" || type === "number") {
        return Number(value) || 0;
    } else if (type === "int") {
        return Math.floor(value) || 0;
    } else {
        return value.toString();
    }
}
