import * as xlsx from "node-xlsx";
import * as fs from "fs";
import * as path from "path";
interface I_in_out {
    "input": string,
    "output": string
}
console.log("\n\n");
let config: { "directory": I_in_out[], "file": I_in_out[] } = require("./config.json");

for (let one of config.directory) {
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
        if (!filename.endsWith(".xlsx")) {
            return;
        }
        let basename = path.basename(filename, '.xlsx');
        let intputfilepath = path.join(one.input, filename);
        let buff = fs.readFileSync(intputfilepath);
        parseBuffToJson(buff, path.join(one.output, basename + ".json"));
        console.log("directory","-----------", filename);
    });
}

console.log("\n");

for (let one of config.file) {
    if (path.extname(one.input) !== ".xlsx") {
        console.error("file", "非法输入路径", one.input);
        continue;
    }
    if (path.extname(one.output) !== ".json") {
        console.error("file", "非法输出路径", one.output);
        continue;
    }
    try {
        fs.readdirSync(path.dirname(one.input));
    } catch (e) {
        console.error("file", "非法输入路径", one.input);
        continue;
    }
    try {
        fs.readdirSync(path.dirname(one.output));
    } catch (e) {
        console.error("file", "非法输出路径", one.output);
        continue;
    }

    let buff: Buffer = null as any;
    try {
        buff = fs.readFileSync(one.input);
    } catch (e) {
        console.error("file", "非法输入路径", one.input);
        continue;
    }
    parseBuffToJson(buff, one.output);
    console.log("file     ","-----------", path.basename(one.input));
}
console.log("\n\n");

function parseBuffToJson(buff: Buffer, outputfile: string) {
    let sheets = xlsx.parse(buff);
    let lists: any = sheets[0].data;
    let keyarr = lists[0];
    let typearr = lists[1];
    let obj: any = {};
    for (let i = 3; i < lists.length; i++) {
        obj[lists[i][0]] = createObj(keyarr, typearr, lists[i]);
    }
    fs.writeFileSync(outputfile, JSON.stringify(obj, null, 4));
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
