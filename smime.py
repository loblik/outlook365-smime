#!/usr/bin/env -S python3 -u

import json
import sys
import os
import traceback

# >>> a['content']['recipient_infos'][1].parse()['rid'].parse()['serial_number'].native
# >>> cms.ContentInfo.load(data)
# from asn1crypto import cms

MS_EXCHANGE_SMIME = ":#Microsoft.Exchange.Clients.BrowserExtension.Smime"

MS_EXCHANGE_CLIENTS_SMIME = "#Microsoft.Exchange.Clients.Smime"

RESULT_SMIME = 'ReturnPartialSmimeResult' + MS_EXCHANGE_SMIME

EXTENSION_MESSAGE = "ExtensionMessage" + MS_EXCHANGE_SMIME
POST_PARTIAL_SMIME_REQUEST = "PostPartialSmimeRequest" + MS_EXCHANGE_SMIME

ACK_PARTIAL_SMIME_REQUEST_ARRIVED = "AcknowledgePartialSmimeRequestArrived" + MS_EXCHANGE_SMIME

INITIALIZE_PARAMS = "InitializeParams" + MS_EXCHANGE_SMIME
SMIME_CONTROL_CAPS = "SmimeControlCapabilities" + MS_EXCHANGE_CLIENTS_SMIME

SMIME_APP_VERSION = "4.0800.20.19.814.2"

def getReqKey(json_data):
    return str(json_data['portId']) + ":" + str(json_data['requestId'])

class SmimeCommands:
    def __init__(self):
        self.handlers = dict()
        self.handlers["InitializeParams" + MS_EXCHANGE_SMIME] = self.handleInitializeParams
        self.handlers["CreateMessageFromSmimeParams" + MS_EXCHANGE_SMIME] = self.createMessageFromSmimeParams

    def handleInitializeParams(self, json_data):
        log("SETTINGS RECEIVED: " + dump_json(json_data))
        if not 'Settings' in json_data:
            raise Exception("No Settings found")

        log("\n\nsettings\n")
        #log(dump_json(json.loads(json_data['Settings'])))

        log("return handle settings\n\n")
        return self.buildSmimeControlsCaps(json_data)

    def createMessageFromSmimeParams(self, json_data):
        log("CREATE FROM SMIME PARAMS\n")
        log(dump_json(json_data))
        log("\nEND\n")
        return

    def buildSmimeControlsCaps(self, json_data):
        json_data = dict()
        json_data['__type'] = SMIME_CONTROL_CAPS
        json_data['SupportsAsyncMethods'] = True
        json_data['Version'] = SMIME_APP_VERSION
        return json_data

    def handleCommand(self, json_data):
        if not '__type' in json_data:
            raise Exception("unspecified __type")

        cmd = json_data['__type']

        if not cmd in self.handlers:
            raise Exception("no handler found for: " + cmd)

        return self.handlers[cmd](json_data)

class Request:
    def __init__(self):
        self.requests = dict()
        self.data = ""
        self.finished = False
        self.offset = 0
        self.part = 0

    def addData(self, data):
        self.data += data['PartialData']
        self.finished = data['IsLastPart']

    def isFinished(self):
        return self.finished

class RequestMap:
    def __init__(self):
        self.requests = dict()
        self.cmds = SmimeCommands()

    def addRequest(self, req_key, request):
        self.requests[req_key] = request
        return request

    def getRequest(self, req_key):
        if not req_key in self.requests:
            return self.addRequest(req_key, Request())

        return self.requests[req_key]

    def buildUploadPartialRequestAck(self):
        json_data = dict()
        json_data['__type'] = ACK_PARTIAL_SMIME_REQUEST_ARRIVED
        json_data['PartIndex'] = -1
        json_data['StartOffset'] = -1
        json_data['NextStartOffset'] = -1
        json_data['Status'] = 1
        return json_data

    def buildDownloadPartialResult(self, index, chunk_size, is_last, data):
        json_data = dict()
        json_data['__type'] = "ReturnPartialSmimeResult" + MS_EXCHANGE_SMIME
        json_data['PartIndex'] = index
        json_data['StartOffset'] = chunk_size * index
        json_data['EndOffset'] = chunk_size * index + len(data)
        json_data['IsLastPart'] = is_last
        json_data['PartialData'] = data
        return json_data

    def handleUpload(self, key, data, sendCb):
        req = self.getRequest(key)

        req.addData(data)

        return sendCb(self.buildUploadPartialRequestAck())

    def handleDownload(self, key, data, sendCb):
        req = self.getRequest(key)

        if not req.isFinished():
            raise Exception("request is not finished yet")

        json_data = json.loads(self.requests[key].data)

        resp_data = self.cmds.handleCommand(json_data)

        json_inner = dict()
        json_inner['ErrorCode'] = 0
        json_inner['Data'] = resp_data

        resp_data = json.dumps(json_inner)

        chunk_size = data["MaxPartSize"]
        log("RESPONSE " + str(len(resp_data)))
        log(str(resp_data))

        rest_bytes = len(resp_data)

        index = 0
        while rest_bytes > 0:
            chunk_data = resp_data[index*chunk_size:chunk_size]
            data = self.buildDownloadPartialResult(index, chunk_size, rest_bytes <= chunk_size, resp_data)

            rest_bytes -= len(chunk_data)
            index += 1

            sendCb(data)

        return


class SmimeApp:
    def __init__(self):
        self.requests = RequestMap()

#    def processRequest(self, request):
    def handlePartialResponse(self, req_port, req_id, msg_type, data):
        response = dict()
        response['data'] = data
        response['messageType'] = msg_type
        response['portId'] = req_port
        response['requestId'] = req_id
        self.sendNativeMsg(json.dumps(response))

    def sendNativeMsg(self, data):
        length = len(data)
        sys.stdout.buffer.write(length.to_bytes(4, sys.byteorder))
        sys.stdout.buffer.write(bytes(data, 'utf-8'))
        sys.stdout.flush()

        log("\n\nMSG-RESPONSE: <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< len=" + str(length) + "\n" + dump_json(json.loads(data)))
        log(str(data))

    def recieveNativeMsg(self, data):
        json_data = json.loads(data)

        if not 'messageType' in json_data:
            raise Exception("missing messageType")

        if not 'data' in json_data:
            raise Exception("missing data")

        if not 'portId' in json_data:
            raise Exception("missing portId")

        if not 'requestId' in json_data:
            raise Exception("missing requestId")

        req_port = json_data['portId']
        req_id = json_data['requestId']
        req_key = str(req_port) + ":" + str(req_id)
        data = json_data['data']
        msg_type = json_data['messageType']

        cb = lambda response: self.handlePartialResponse(req_port, req_id, msg_type, response)

        if msg_type == "UploadPartialRequest":
            self.requests.handleUpload(req_key, data, cb)
        elif msg_type  == 'DownloadPartialResult':
            self.requests.handleDownload(req_key, data, cb)
        elif msg_type  == 'GetSettings':
            self.sendNativeMsg('{"AllowedDomainsByPolicy":[]}')
        else:
            raise Exception("unkown messageType " + msg_type)

    def run(self):
        while True:
            # native byte oder
            length_bytes  = sys.stdin.buffer.read(4)
            length = int.from_bytes(length_bytes, sys.byteorder)

            if length == 0:
                break

            data = sys.stdin.buffer.read(length)

            log("\n\nNEW-MSG: >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> len=" + str(length))
            log("\n-----\n")
            log(str(data))
            log("\n-----\n")
            log(dump_json(json.loads(data)))
            log("\n=====\n")

            app.recieveNativeMsg(data)

def dump_json(json_data):
    return json.dumps(json_data, sort_keys=True, indent=4)

def log(data):
    logh.write(data)
    logh.flush()
    sys.stderr.write(data)
    sys.stderr.write('\n')

logh = open("/tmp/smime.log","a+")

try:
    app = SmimeApp()
    app.run()

except Exception as e:
    logh.write(str(e))
    logh.write(traceback.format_exc())
