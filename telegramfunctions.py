import asyncio
import telegram

async def sendTelegramMsg(msg,chat_id):
    bot = telegram.Bot("6557815543:AAFYQ1itku7fMSbBFirk3YQIQRtNE54WbEM")
    async with bot:
        await bot.send_message(text=msg, chat_id=chat_id)
        
async def sendTelegramMsgWithDocuments(chat_id):
    bot = telegram.Bot("6557815543:AAFYQ1itku7fMSbBFirk3YQIQRtNE54WbEM")
    async with bot:        
        await bot.send_document(chat_id=chat_id, document='C:\\BOT INSUMO BASE\\Input BOT 001 V1.xlsx')       
