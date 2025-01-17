// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Colors, ConfigFolderName, LogLevel, LogProvider } from "@microsoft/teamsfx-api";
import chalk from "chalk";
import * as os from "os";
import * as path from "path";
import { SuccessText, TextType, WarningText, colorize, replaceTemplateString } from "../colorize";
import ScreenManager from "../console/screen";
import { CLILogLevel } from "../constants";
import { strings } from "../resource";
import { getColorizedString } from "../utils";

export class CLILogProvider implements LogProvider {
  private static instance: CLILogProvider;

  private static logLevel: CLILogLevel = CLILogLevel.error;

  private logFileName: string;

  constructor() {
    this.logFileName = `${new Date().toISOString().replace(/-|:|\.\d+Z$/g, "")}.log`;
  }

  public getLogLevel() {
    return CLILogProvider.logLevel;
  }

  public setLogLevel(logLevel: CLILogLevel) {
    CLILogProvider.logLevel = logLevel;
  }

  /**
   * Gets instance
   * @returns instance
   */
  public static getInstance(): CLILogProvider {
    if (!CLILogProvider.instance) {
      CLILogProvider.instance = new CLILogProvider();
    }

    return CLILogProvider.instance;
  }

  getLogFilePath(): string {
    return path.join(os.tmpdir(), `.${ConfigFolderName}`, "cli-log", this.logFileName);
  }

  trace(message: string): Promise<boolean> {
    return this.log(LogLevel.Trace, message);
  }

  debug(message: string): Promise<boolean> {
    return this.log(LogLevel.Debug, message);
  }

  info(message: Array<{ content: string; color: Colors }>): Promise<boolean>;

  info(message: string): Promise<boolean>;

  info(message: string | Array<{ content: string; color: Colors }>): Promise<boolean> {
    if (message instanceof Array) {
      message = getColorizedString(message);
    } else {
      message = chalk.whiteBright(message);
    }
    return this.log(LogLevel.Info, message);
  }

  white(msg: string): string {
    return chalk.whiteBright(msg);
  }

  warning(message: string): Promise<boolean> {
    return this.log(LogLevel.Warning, message);
  }

  error(message: string): Promise<boolean> {
    return this.log(LogLevel.Error, message);
  }

  fatal(message: string): Promise<boolean> {
    return this.log(LogLevel.Fatal, message);
  }

  linkColor(msg: string): string {
    return chalk.cyanBright.underline(msg);
  }

  async log(logLevel: LogLevel, message: string): Promise<boolean> {
    switch (logLevel) {
      case LogLevel.Trace:
      case LogLevel.Debug:
        if (CLILogProvider.logLevel === CLILogLevel.debug) {
          this.outputDetails(message);
        }
        break;
      case LogLevel.Info:
        if (
          CLILogProvider.logLevel === CLILogLevel.debug ||
          CLILogProvider.logLevel === CLILogLevel.verbose
        ) {
          this.outputDetails(message);
        }
        break;
      case LogLevel.Warning:
        if (CLILogProvider.logLevel !== CLILogLevel.error) {
          this.outputWarning(message);
        }
        break;
      case LogLevel.Error:
      case LogLevel.Fatal:
        this.outputError(message);
        break;
    }
    return true;
  }

  outputSuccess(template: string, ...args: string[]): void {
    ScreenManager.writeLine(
      SuccessText + colorize(replaceTemplateString(template, ...args), TextType.Info)
    );
  }

  outputInfo(template: string, ...args: string[]): void {
    ScreenManager.writeLine(colorize(replaceTemplateString(template, ...args), TextType.Info));
  }

  outputDetails(template: string, ...args: string[]): void {
    ScreenManager.writeLine(colorize(replaceTemplateString(template, ...args), TextType.Details));
  }

  outputWarning(template: string, ...args: string[]): void {
    ScreenManager.writeLine(
      WarningText + colorize(replaceTemplateString(template, ...args), TextType.Info),
      true
    );
  }

  outputError(template: string, ...args: string[]): void {
    ScreenManager.writeLine(
      colorize(strings["error.prefix"] + replaceTemplateString(template, ...args), TextType.Error),
      true
    );
  }

  necessaryLog(logLevel: LogLevel, message: string, white = false) {
    switch (logLevel) {
      case LogLevel.Trace:
      case LogLevel.Debug:
        this.outputDetails(message);
        break;
      case LogLevel.Info:
        if (white) {
          this.outputInfo(message);
        } else {
          ScreenManager.writeLine(chalk.greenBright(message));
        }
        break;
      case LogLevel.Warning:
        this.outputWarning(message);
        break;
      case LogLevel.Error:
      case LogLevel.Fatal:
        this.outputError(message);
        break;
    }
  }

  rawLog(message: string) {
    process.stdout.write(message);
  }
}

export default CLILogProvider.getInstance();
