module.exports = async function (context, req) {
    context.log('JavaScript HTTP trigger function processed a request.');

    // JSON response
    context.res = {
        status: 200,
        headers: {
            'Content-Type': 'application/json'
        },
        body: {
            msg: `Hello, world! - ${new Date().toISOString()}`
        }
    };
};
