#!/usr/bin/env node
import { runCli } from "./cli-runner.js";

void runCli(process.argv.slice(2)).then((exitCode) => {
  process.exitCode = exitCode;
});
