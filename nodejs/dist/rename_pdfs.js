"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
const fs = __importStar(require("fs"));
const path = __importStar(require("path"));
const outputDir = path.join(__dirname, '..', 'output');
const mappingFilePath = path.join(outputDir, 'filename_mapping.json');
// Read the mapping file
fs.readFile(mappingFilePath, 'utf8', (err, data) => {
    if (err) {
        console.error('Error reading mapping file:', err);
        return;
    }
    try {
        const mappings = JSON.parse(data);
        // Iterate through the mappings and rename files
        for (const generatedName in mappings) {
            if (mappings.hasOwnProperty(generatedName)) {
                const originalName = mappings[generatedName].originalName;
                const oldPath = path.join(outputDir, generatedName);
                const newPath = path.join(outputDir, `${originalName}.pdf`);
                fs.rename(oldPath, newPath, (renameErr) => {
                    if (renameErr) {
                        console.error(`Error renaming file ${generatedName}:`, renameErr);
                    }
                    else {
                        console.log(`Renamed ${generatedName} to ${originalName}.pdf`);
                    }
                });
            }
        }
    }
    catch (parseError) {
        console.error('Error parsing mapping file:', parseError);
    }
});
