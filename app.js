/*-----------------------------------------------------------------------------
A simple echo bot for the Microsoft Bot Framework. 
-----------------------------------------------------------------------------*/

var restify = require('restify');
var builder = require('botbuilder');
var request = require('request').defaults({ encoding: null });

var luis = require('./luis.json');

// Setup Restify Server
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
   console.log('%s listening to %s', server.name, server.url); 
});
  
// Create chat connector for communicating with the Bot Framework Service
var connector = new builder.ChatConnector({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword,
    stateEndpoint: process.env.BotStateEndpoint,
    openIdMetadata: process.env.BotOpenIdMetadata 
});

// Listen for messages from users 
server.post('/api/messages', connector.listen());

/*----------------------------------------------------------------------------------------
* Bot Storage: This is a great spot to register the private state storage for your bot. 
* We provide adapters for Azure Table, CosmosDb, SQL Azure, or you can implement your own!
* For samples and documentation, see: https://github.com/Microsoft/BotBuilder-Azure
* ---------------------------------------------------------------------------------------- */

// Create your bot with a function to receive messages from the user
var bot = new builder.UniversalBot(connector, function (session) {
    session.send('Sorry, the OMS autobot didn\'t understand \'%s\'. Type \'order\' if you would like to place an order.', session.message.text);
});

var botLoggerHostName =  process.env.BotLogHostName; //'http://localhost:8080'

// Make sure you add code to validate these fields
var luisAppId = process.env.LuisAppId;
var luisAPIKey = process.env.LuisAPIKey;
var luisAPIHostName = process.env.LuisAPIHostName || 'westus.api.cognitive.microsoft.com';

const LuisModelUrl = 'https://' + luisAPIHostName + '/luis/v2.0/apps/' + luisAppId + '?subscription-key=' + luisAPIKey + '&verbose=true&timezoneOffset=0&q=';

console.log(LuisModelUrl)

var recognizer = new builder.LuisRecognizer(LuisModelUrl);
bot.recognizer(recognizer);

// Order Query Functionality
bot.dialog('OrderQuery', [
    function (session, args, next) {
        var intent = args.intent
        console.log(intent.entities)
        var stock = builder.EntityRecognizer.findEntity(intent.entities, 'Stocks')

        if (stock && stock.resolution){
            console.log(stock.entity)
            getHoldingForStock(session.message.address,stock.entity.toString())
        } else {
            getHoldings(session.message.address)

        }
        session.endDialog()
    }
]).triggerAction({
    matches: 'OrderQuery',
    /*TODO:: disable confirmation prompt to avoid 'ibm'/'microsoft' stock confirmation triggering unwanted new dialog confirmation*/
    //confirmPrompt: "This will cancel the creation of order you started. Are you sure?"
}).cancelAction('cancelCreateNote', "Query canceled.", {
    matches: /^(cancel|nevermind)/i,
    confirmPrompt: "Are you sure you want to make a query request?"
});

// ORDER Functionality
bot.dialog('Order', [
    function (session, args, next) {
        var order = session.dialogData.order = {}

        if (args && args.isReprompt && args.dialogData && args.dialogData.order){
            // We were sent here to revalidate some user input
            // Reinitialise the order data using the args
            order = session.dialogData.order = args.dialogData.order
        }
        else if (args && args.intent && args.intent.entities){
            // LUIS recognizer triggered dialog
            // Scrape out all intent entities
           // var {intent} = args
            var intent = args.intent
            console.log(intent.entities)

            var stock = builder.EntityRecognizer.findEntity(intent.entities, 'Stocks')
            if (stock && stock.resolution){
                order.stock = stock.resolution.values[0]
            }

            var quantity = builder.EntityRecognizer.findEntity(intent.entities, 'builtin.number')
            if (quantity && quantity.resolution){
                order.qty = quantity.resolution.value
            }

            var direction = builder.EntityRecognizer.findEntity(intent.entities, 'OrderDirection')

            if (direction && direction.entity){
                order.direction = getNormalisedDirection(direction.entity.toLowerCase())

            }
        }

        if (!order.stock) {
            promptForText(session, order, 'What stock would you like to order?')
        } else {
            next()
        }

    },
    /*
        --1-- Validate stock
     */
    function (session, results, next) {

        console.log(session.dialogData)
        //var {dialogData} = session
        var dialogData = session.dialogData
        var {order} = dialogData

        if (!order.stock) {
            // Example response: { index: 0, entity: 'Apple', score: 0.8 }
            const bestMatch = builder.EntityRecognizer.findBestMatch(getStockListFromLuisConfig(), results.response, 0.6)
            console.log(bestMatch)
            if (bestMatch){
                order.stock = bestMatch.entity
            }
        }

        if (!order.stock) {
            // Unable to validate stock, send back to the start...
            session.replaceDialog('Order', {dialogData: dialogData, isReprompt: true});
        }
        else if (!order.qty){
            promptForNumber(session, order, 'How many '+order.stock+' would you like to order?')
        } else {
            next()
        }

    },
    /*
        --2-- Validate qty
     */
    function (session, results, next) {
        var {dialogData} = session
        var {order} = dialogData

        if (!order.qty) {
            order.qty=results.response
        }

        if (!order.direction){
            promptForChoice(session, order, 'Would you like to buy or sell '+order.stock+'?', ['Buy','Sell'], builder.ListStyle.button)
        } else {
            next()
        }

    },
    /*
        --3-- Validate direction
     */
    function (session, results) {
        var {dialogData} = session
        var {order} = dialogData

        if (!order.direction) {
            // check the confirmation response
            console.log(results.response)
            if (results.response.entity){
                order.direction=results.response.entity
            }
            else {
                console.log('Something went wrong, I shouldn\'t end up here')
            }

        }
        var price = getSharePrice(order.stock).toFixed(2)
        order.price = price
        var totalcost = (price * order.qty)
        console.log(order)
        promptForConfirmation(session, order, 'Confirm you would like to place a '+order.direction+' order for '+niceNumber(order.qty)+' of '+order.stock+' at $'+order.price+' AUD? The total cost is $'+niceNumber(totalcost));
    },
    /*
        --4-- Confirmation?
     */
    function (session, results) {
        var {dialogData} = session
        var {order} = dialogData
        var totalcost = (order.price * order.qty)
        order.completed=results.response
        promptForText(session, order, order.completed?'OK, order completed for a total value of $'+niceNumber(totalcost)+' AUD at an average price of $'+order.price+' per share.':'Order cancelled.')
        session.endDialog()
    }
]).triggerAction({
    matches: 'Order',
    /*TODO:: disable confirmation prompt to avoid 'ibm'/'microsoft' stock confirmation triggering unwanted new dialog confirmation*/
    //confirmPrompt: "This will cancel the creation of order you started. Are you sure?"
}).cancelAction('cancelCreateNote', "Order canceled.", {
    matches: /^(cancel|nevermind)/i,
    confirmPrompt: "Are you sure you want to stop ordering?"
});


const promptForOrderQueryDetails = (session, text) => {
    builder.Prompts.confirm(session, text, {
        speak: text,
        retrySpeak: text+' Say cancel to dismiss me',
        inputHint: builder.InputHint.expectingInput,
    })
 //   logOrderState(session, order, text)
}


const promptForConfirmation = (session, order, text) => {
    builder.Prompts.confirm(session, text, {
        speak: text,
        retrySpeak: text+' Say cancel to dismiss me',
        inputHint: builder.InputHint.expectingInput,
    })

    logOrderState(session, order, text)
}


const promptForChoice = (session, order, text, choices, listStyle) => {
    builder.Prompts.choice(session, text, choices, {
        listStyle:listStyle,
        speak: text,
        retrySpeak: text+' Say cancel to dismiss me',
        inputHint: builder.InputHint.expectingInput
    })
    logOrderState(session, order, text, choices)
}

const promptForText = (session, order, text) => {
    builder.Prompts.text(session, text, {
        speak: text,
        retrySpeak: text+' Say cancel to dismiss me',
        inputHint: builder.InputHint.expectingInput,
    })
    logOrderState(session, order, text)
}

const promptForNumber = (session, order, text) => {
    builder.Prompts.number(session, text, {
        speak: text,
        retrySpeak: text+' Say cancel to dismiss me',
        inputHint: builder.InputHint.expectingInput,
    })
    logOrderState(session, order, text)
}



const logOrderState = (session, order, message, choices) => {

    var conversationId, channel, lastUserMessage, lastOrderState, lastSystemMessage, userName, userId, timestamp
    if (session.message && session.message.address && session.message.address.conversation){
        //console.log('Conversationid:')
        conversationId = session.message.address.conversation.id
        channel = session.message.address.channelId
        timestamp = session.message.timestamp
    }
    if (session.message && session.message.type === 'message'){
        //console.log('User message:')
        lastUserMessage = session.message.text
    }
    if (session.message && session.message.user){
        userId = session.message.user.id
        userName = session.message.user.name
    }
    if (order){
        lastOrderState = order
        lastOrderState.timestamp=timestamp
    }
    if (message){
        //console.log('System message:')
        lastSystemMessage = message
    }

    const data = {
        conversationId:conversationId,
        channel:channel,
        lastUserMessage:lastUserMessage,
        lastOrderState:lastOrderState,
        lastSystemMessage:lastSystemMessage,
        choices:choices?choices:[],
        userId:userId,
        userName:userName,
        timestamp:timestamp
    }
    console.log(data)
    performOrderStateLogging(data)
}

const performOrderStateLogging = (data) => {
    var logUrl=botLoggerHostName+'/order/log'
    console.log('logging order state to url: '+logUrl)
    var requestData = {
        url: logUrl,
        body: data,
        json: true
    };
    request.put(requestData, function (error, response, body) {})
};

function getHoldings(address)  {
    var getHoldingsUrl=botLoggerHostName+'/holdings'
    console.log('Getting Holdings: '+getHoldingsUrl)
    var requestData = {
        url: getHoldingsUrl,
        json: true
    };
    request.get(requestData, function (error, response, body) {
        console.log(body)
        if (body&& !error){
            var text
            var stocks = Object.keys(body);
            if(!stocks || stocks.length === 0){
                text = "Ok you have not placed any orders yet. Say Orders to get started."
            } else {
                text = "Ok, you have"+stocks.map((stock)=>' '+body[stock]+' shares of '+stock+'')
            }

            var msg = new builder.Message().address(address)
            msg.text(text)
            msg.textLocale('en-US')
            bot.send(msg)
        }
    })

   // return body
    //request.put(requestData, function (error, response, body) {})
};

function getHoldingForStock(address, stock)  {
    var getHoldingForStockUrl=botLoggerHostName+'/stock/holding'
    console.log('Getting Holdings: '+getHoldingForStockUrl)
    var requestData = {
        url: getHoldingForStockUrl,
        qs: {stock:stock},
        json: true
    };

    request.get(requestData, function (error, response, body) {
        console.log(body)
        if (body && !error){
            var msg = new builder.Message().address(address)
            if(body && body.qty){
                var avgPrice = body.avgPrice;
                msg.text("Ok, you have "+body.qty+' shares of '+stock+' at an average price of $'+(avgPrice * 1).toFixed(2))
            } else {
                msg.text("Ok, you have no shares of "+stock+"!")
            }
            msg.textLocale('en-US')
            bot.send(msg)
        }
    })

};

/*
* "closedLists": [
    {
      "name": "Stocks",
      "subLists": [
        {
          "canonicalForm": "Apple",
          "list": []
        },
        {
          "canonicalForm": "Optus",
          "list": []
        },
        ...
      ]
    }*/
function getStockListFromLuisConfig() {
    const stockList = luis.closedLists.filter(list=>list.name === 'Stocks')[0]
    return stockList.subLists.map(element=>element.canonicalForm)
}


// Function to determine the Order direction based on various definitions
function getNormalisedDirection(directionStr) {
    if (directionStr.toLowerCase() === 'sell' || directionStr.toLowerCase() === 'loose' || directionStr.toLowerCase() === 'short' || directionStr.toLowerCase() === 'cut' ){
        return 'sell'
    }
    return 'buy'
}


// Function to determine the Order direction based on various definitions
function getSharePrice(stock) {
    var min = 10
    var max = 24
    if (stock.toLowerCase()==='ibm'){
        min = 170
        max = 210
    } else if(stock.toLowerCase() ==='microsoft'){
        min = 80
        max = 101
    } else if(stock.toLowerCase() ==='apple') {
        min = 160
        max = 190
    } else if(stock.toLowerCase() ==='sony') {
        min = 30
        max = 40
    }

    return getRandomPrice(min, max);
}

function getRandomPrice(min, max) {
    return Math.random() * (max - min) + min;
}

function niceNumber(num){
    return num.toLocaleString(undefined);
}

function niceDecimal(num){
    return num.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2});
}
