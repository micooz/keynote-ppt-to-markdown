#!/usr/bin/env node
const { runCli } = require('../dist/index.js');

(async () => {
  try {
    await runCli();
  } catch (error) {
    if (error instanceof Error) {
      console.error(`错误: ${error.message}`);
    } else {
      console.error('发生未知错误:', error);
    }
    process.exit(1);
  }
})();
