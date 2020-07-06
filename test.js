"use strict";
exports.__esModule = true;
var xlsx = require("node-xlsx");
var fs = require("fs");
var path = require("path");
console.log("\n");
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
        if (filename[0] === "~") {
            return;
        }
        if (!filename.endsWith(".xlsx")) {
            return;
        }
        var intputfilepath = path.join(one.input, filename);
        var buff = fs.readFileSync(intputfilepath);
        parseBuffToJson(buff, one.output, path.basename(filename, '.xlsx'));
        console.log("---->>>", one.input, " ", filename);
    });
};
for (var _i = 0, config_1 = config; _i < config_1.length; _i++) {
    var one = config_1[_i];
    _loop_1(one);
}
console.log("\n");
function parseBuffToJson(buff, outputDir, filename) {
    var sheets = xlsx.parse(buff);
    for (var _i = 0, sheets_1 = sheets; _i < sheets_1.length; _i++) {
        var one = sheets_1[_i];
        var lists = one.data;
        if (lists.length === 0) {
            continue;
        }
        var keyarr = lists[0];
        var typearr = lists[1];
        var obj = {};
        for (var i = 3; i < lists.length; i++) {
            if (lists[i][0] === undefined) {
                continue;
            }
            obj[lists[i][0]] = createObj(keyarr, typearr, lists[i]);
        }
        fs.writeFileSync(path.join(outputDir, filename + "_" + one.name + ".json"), JSON.stringify(obj, null, 4));
    }
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
