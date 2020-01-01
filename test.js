"use strict";
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (Object.hasOwnProperty.call(mod, k)) result[k] = mod[k];
    result["default"] = mod;
    return result;
};
Object.defineProperty(exports, "__esModule", { value: true });
var xlsx = __importStar(require("node-xlsx"));
var fs = __importStar(require("fs"));
var path = __importStar(require("path"));
console.log("\n\n");
var config = require("./config.json");
var _loop_1 = function (one) {
    var inputfiles = [];
    try {
        inputfiles = fs.readdirSync(one.input);
    }
    catch (e) {
        console.error("directory", "非法输入路径", one.input);
        return "continue";
    }
    try {
        fs.readdirSync(one.output);
    }
    catch (e) {
        console.error("directory", "非法输出路径", one.output);
        return "continue";
    }
    inputfiles.forEach(function (filename) {
        if (!filename.endsWith(".xlsx")) {
            return;
        }
        var basename = path.basename(filename, '.xlsx');
        var intputfilepath = path.join(one.input, filename);
        var buff = fs.readFileSync(intputfilepath);
        parseBuffToJson(buff, path.join(one.output, basename + ".json"));
        console.log("directory", "-----------", filename);
    });
};
for (var _i = 0, _a = config.directory; _i < _a.length; _i++) {
    var one = _a[_i];
    _loop_1(one);
}
console.log("\n");
for (var _b = 0, _c = config.file; _b < _c.length; _b++) {
    var one = _c[_b];
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
    }
    catch (e) {
        console.error("file", "非法输入路径", one.input);
        continue;
    }
    try {
        fs.readdirSync(path.dirname(one.output));
    }
    catch (e) {
        console.error("file", "非法输出路径", one.output);
        continue;
    }
    var buff = null;
    try {
        buff = fs.readFileSync(one.input);
    }
    catch (e) {
        console.error("file", "非法输入路径", one.input);
        continue;
    }
    parseBuffToJson(buff, one.output);
    console.log("file     ", "-----------", path.basename(one.input));
}
console.log("\n\n");
function parseBuffToJson(buff, outputfile) {
    var sheets = xlsx.parse(buff);
    var lists = sheets[0].data;
    var keyarr = lists[0];
    var typearr = lists[1];
    var obj = {};
    for (var i = 3; i < lists.length; i++) {
        obj[lists[i][0]] = createObj(keyarr, typearr, lists[i]);
    }
    fs.writeFileSync(outputfile, JSON.stringify(obj, null, 4));
}
function createObj(keyarr, typearr, dataarr) {
    var obj = {};
    for (var i = 0; i < keyarr.length; i++) {
        obj[keyarr[i]] = changeValue(dataarr[i], typearr[i]);
    }
    return obj;
}
function changeValue(value, type) {
    if (value === undefined) {
        value = "";
    }
    if (type === "bool") {
        if (typeof value === "string") {
            value = value.trim();
        }
        if (value === 0 || value == "" || value === "false") {
            return false;
        }
        else {
            return true;
        }
    }
    else if (type === "string") {
        return value.toString();
    }
    else if (type === "float" || type === "number") {
        return Number(value) || 0;
    }
    else if (type === "int") {
        return Math.floor(value) || 0;
    }
    else {
        return value.toString();
    }
}
