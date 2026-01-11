"use strict";
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.ReplacementModule = void 0;
const common_1 = require("@nestjs/common");
const replacement_controller_1 = require("./replacement.controller");
const replacement_service_1 = require("./replacement.service");
let ReplacementModule = class ReplacementModule {
};
exports.ReplacementModule = ReplacementModule;
exports.ReplacementModule = ReplacementModule = __decorate([
    (0, common_1.Module)({
        controllers: [replacement_controller_1.ReplacementController],
        providers: [replacement_service_1.ReplacementService],
        exports: [replacement_service_1.ReplacementService],
    })
], ReplacementModule);
//# sourceMappingURL=replacement.module.js.map