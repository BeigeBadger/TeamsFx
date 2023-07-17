// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author yuqzho@microsoft.com
 */

import {
  Context,
  FxError,
  OpenAIManifestAuthType,
  OpenAIPluginManifest,
  Result,
  UserError,
  err,
  ok,
  TeamsAppManifest,
} from "@microsoft/teamsfx-api";
import axios, { AxiosResponse } from "axios";
import { sendRequestWithRetry } from "../utils";
import {
  ErrorType as ApiSpecErrorType,
  ValidationStatus,
} from "../../../common/spec-parser/interfaces";
import { SpecParser } from "../../../common/spec-parser/specParser";
import fs from "fs-extra";
import { manifestUtils } from "../../driver/teamsApp/utils/ManifestUtils";
import path from "path";
import { getLocalizedString } from "../../../common/localizeUtils";

const manifestFilePath = "/.well-known/ai-plugin.json";
const teamsFxEnv = "${{TEAMSFX_ENV}}";
const componentName = "OpenAIPluginManifestHelper";

enum OpenAIPluginManifestErrorType {
  AuthNotSupported,
  ApiUrlMissing,
}

export interface ErrorResult {
  /**
   * The type of error.
   */
  type: ApiSpecErrorType | OpenAIPluginManifestErrorType;

  /**
   * The content of the error.
   */
  content: string;
}

export class OpenAIPluginManifestHelper {
  static async loadOpenAIPluginManifest(domain: string): Promise<OpenAIPluginManifest> {
    const path = domain + manifestFilePath;

    try {
      const res: AxiosResponse<any> = await sendRequestWithRetry(async () => {
        return await axios.get(path);
      }, 3);

      return res.data;
    } catch (e) {
      throw new UserError(
        componentName,
        "loadOpenAIPluginManifest",
        getLocalizedString("error.copilotPlugin.openAiPluginManifest.CannotGetManifest", path),
        getLocalizedString("error.copilotPlugin.openAiPluginManifest.CannotGetManifest", path)
      );
    }
  }

  static async updateManifest(
    openAiPluginManifest: OpenAIPluginManifest,
    teamsAppManifest: TeamsAppManifest,
    manifestPath: string
  ): Promise<Result<undefined, FxError>> {
    teamsAppManifest.name.full = openAiPluginManifest.name_for_model;
    teamsAppManifest.name.short = `${openAiPluginManifest.name_for_human}-${teamsFxEnv}`;
    teamsAppManifest.description.full = openAiPluginManifest.description_for_model;
    teamsAppManifest.description.short = openAiPluginManifest.description_for_human;
    teamsAppManifest.developer.websiteUrl = openAiPluginManifest.legal_info_url;
    teamsAppManifest.developer.privacyUrl = openAiPluginManifest.legal_info_url;
    teamsAppManifest.developer.termsOfUseUrl = openAiPluginManifest.legal_info_url;

    await fs.writeFile(manifestPath, JSON.stringify(teamsAppManifest, null, "\t"), "utf-8");
    return ok(undefined);
  }
}

export async function listOperations(
  context: Context,
  manifest: OpenAIPluginManifest | undefined,
  apiSpecUrl: string | undefined,
  shouldLogWarning = true
): Promise<Result<string[], ErrorResult[]>> {
  if (manifest) {
    apiSpecUrl = manifest.api.url;
    const errors = validateOpenAIPluginManifest(manifest);
    if (errors.length > 0) {
      return err(errors);
    }
  }

  const specParser = new SpecParser(apiSpecUrl!);
  const validationRes = await specParser.validate();

  if (validationRes.status === ValidationStatus.Error) {
    return err(validationRes.errors);
  }

  if (shouldLogWarning && validationRes.warnings.length > 0) {
    for (const warning of validationRes.warnings) {
      context.logProvider.warning(warning.content);
    }
  }

  const operations = await specParser.list();
  return ok(operations);
}

function validateOpenAIPluginManifest(manifest: OpenAIPluginManifest): ErrorResult[] {
  const errors: ErrorResult[] = [];
  if (!manifest.api.url) {
    errors.push({
      type: OpenAIPluginManifestErrorType.ApiUrlMissing,
      content: "Missing url in manifest",
    });
  }

  if (manifest.auth.type !== OpenAIManifestAuthType.None) {
    errors.push({
      type: OpenAIPluginManifestErrorType.AuthNotSupported,
      content: "Auth type not supported",
    });
  }
  return errors;
}

function validateTeamsManifestLength(teamsManifest: TeamsAppManifest) {
  const nameShortLimit = 30;
  const nameFullLimit = 100;
  const descriptionShortLimit = 80;
  const descriptionFullLimit = 4000;
  // validate name
  if (teamsManifest.name.short.length === 0) {
    // (×) Error: Short name of the app cannot be empty.
  }

  if (teamsManifest.name.short.length > nameShortLimit) {
    // (×) Error: /name/short must NOT have more than 30 characters
  }

  if (!teamsManifest.name.full?.length) {
    // (×) Error: Short name of the app cannot be empty.  ?? (not required by may be required for Copilot plugin)
  }

  if (teamsManifest.name.full!.length > nameFullLimit) {
    // /name/full must NOT have more than 100 characters
  }

  // validate description

  if (teamsManifest.description.short.length === 0) {
    // (×) Error: Short Description can not be empty
    // Learn more: https://docs.microsoft.com/en-us/microsoftteams/platform/resources/schema/manifest-schema#description
  }
  if (teamsManifest.description.short.length > descriptionShortLimit) {
    // (×) Error: /description/short must NOT have more than 80 characters
  }
  if (teamsManifest.description.full?.length === 0) {
    // (×) Error: Full Description cannot be empty.
  }
  if (teamsManifest.description.full!.length > descriptionFullLimit) {
    // (×) Error: /description/full must NOT have more than 4000 characters
  }
}
