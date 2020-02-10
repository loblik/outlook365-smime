#!/usr/bin/env python3

import json
import sys
import os

MS_SMIME_MSG = "ExtensionMessage:#Microsoft.Exchange.Clients.BrowserExtension.Smime"
MS_SMIME_MSG_PARTIAL_REQ = "PostPartialSmimeRequest:#Microsoft.Exchange.Clients.BrowserExtension.Smime"
MS_SMIME_MSG_PARAMS = "InitializeParams:#Microsoft.Exchange.Clients.BrowserExtension.Smime"

def dump_json(json_data):
    print(json.dumps(json_data, sort_keys=True, indent=4))

def process_smime_msg_params(json_data):
    dump_json(json.loads(json_data["Settings"]))

def process_upload_partial_request(json_data):
    data = json_data["PartialData"]

    json_data = json.loads(data)

    if not '__type' in json_data:
        # TODO log invalid message
        return

    dump_json(json_data)
    print("dasdsadas")

    if json_data['__type'] == MS_SMIME_MSG_PARAMS:
        process_smime_msg_params(json_data)

def process_smime_msg(json_data):
    if not 'messageType' in json_data:
        # TODO log invalid message
        return

    if json_data['messageType'] == "UploadPartialRequest":
        process_upload_partial_request(json_data['data'])

def process_msg(data):
    json_data = json.loads(data)

    dump_json(json_data)

    if not '__type' in json_data:
        # TODO log invalid message
        return

    if json_data['__type'] == MS_SMIME_MSG:
        process_smime_msg(json_data)


while True:
    # native byte oder
    length_bytes  = sys.stdin.buffer.read(4)
    length = int.from_bytes(length_bytes, sys.byteorder)

    if length == 0:
        break

    print("NEW-MSG: len=" + str(length))
    data = sys.stdin.buffer.read(length)

    process_msg(data)
