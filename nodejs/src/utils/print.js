const chalk = require('chalk');

function printInfo(text) {
    console.log(`ℹ️ ${text}`);
}

function printError(message) {
	console.error(chalk.red(`❌ ${message}`));
    process.exitCode = 1;
}

const consoleError = (message) => {
    console.log(message);
}
module.exports = {
	printInfo,
    printError,
    consoleError
};