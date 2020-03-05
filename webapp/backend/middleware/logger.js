module.exports = logger;

function logger(req, res, next) {
	console.log(`${req.method} to ${req.url} at ${new Date().toISOString()}`);
	next();
}
