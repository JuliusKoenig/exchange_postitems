from distutils.dir_util import copy_tree

import urllib3
import argparse
import json
import logging
import os
import sys
import exchangelib

name = "exchange_postitems"
description = "This script allows you to publish PostItems to the public folders of a Microsoft Exchange Server."
version = "v0.1"

sample_config = {
    "server": "127.0.0.1",
    "username": "test.local\\example",
    "password": "",
    "primary_smtp_address": "exampleg@test.local",
    "ssl_verify": True,
    "target_path": "",
    "author": "TEST",
    "subject": "This is a test."
}


def get_folder(folder, root_folder):
    split_folder = folder.split("/")
    c_folder = split_folder.pop(0)
    for f in root_folder.children:
        if f.name == c_folder:
            if len(split_folder) > 0:
                next_folder = "".join(split_folder)
                return get_folder(next_folder, f)
            else:
                return f


def create_new_config(path):
    f = open(path, "w", encoding='utf-8')
    json.dump(sample_config, f)
    f.close()
    return sample_config


if __name__ == '__main__':
    if hasattr(sys, "_MEIPASS"):
        copy_tree(sys._MEIPASS + "\\pytz", sys._MEIPASS + "\\tzdata")

    # parse arguments
    main_parser = argparse.ArgumentParser(description=name + ' - ' + description)
    main_parser.add_argument('-V', '--version', action='version', version=name + " - " + version)
    main_parser.add_argument('-v', '--verbose', action='store_true', help='Show debugging information.')

    sub_parsers = main_parser.add_subparsers(dest='mode')

    mode_auto_parser = sub_parsers.add_parser('mode_auto', help='Automatic config driven mode.')
    mode_auto_parser.add_argument("-c", "--config_file", required=True,
                                  help='A file with all the necessary configuration parameters. '
                                       'If it does not exist, it will be created during the next run.')
    mode_auto_parser.add_argument("--author", help="the author in new post")
    mode_auto_parser.add_argument("--subject", help="the subject in new post")

    mode_manual_parser = sub_parsers.add_parser('mode_manual', help='Manual mode.')
    mode_manual_parser.add_argument("--server", required=True, help="MS Exchangeserver")
    mode_manual_parser.add_argument("--username", required=True, help="the logon username")
    mode_manual_parser.add_argument("--password", required=True, help="the logon password")
    mode_manual_parser.add_argument("--primary_smtp_address", required=True,
                                    help="the primary smtp address for this account.")
    mode_manual_parser.add_argument("--ssl_verify", default=sample_config["ssl_verify"], action='store_false',
                                    help="If this parameter is set, SSL verification is skipped.")
    mode_manual_parser.add_argument("--target_path", required=True, help="the target path in public folders")
    mode_manual_parser.add_argument("--author", required=True, help="the author in new post")
    mode_manual_parser.add_argument("--subject", required=True, help="the subject in new post")

    cmd_args = main_parser.parse_args()

    # check if verbose flag set and set log lvl
    if cmd_args.verbose:
        log_lvl = logging.DEBUG
    else:
        log_lvl = logging.WARNING
    logging.basicConfig(level=log_lvl)

    if cmd_args.mode == "mode_auto":
        if os.path.isfile(cmd_args.config_file):
            with open(cmd_args.config_file, encoding='utf-8') as fh:
                config = json.load(fh)
        else:
            create_new_config(cmd_args.config_file)
            logging.info("Config file created.")
            sys.exit(0)

        if cmd_args.author is not None:
            config["author"] = cmd_args.author
        if cmd_args.subject is not None:
            config["subject"] = cmd_args.subject
    elif cmd_args.mode == "mode_manual":
        config = {"server": cmd_args.server, "username": cmd_args.username, "password": cmd_args.password,
                  "primary_smtp_address": cmd_args.primary_smtp_address, "ssl_verify": cmd_args.ssl_verify,
                  "target_path": cmd_args.target_path, "author": cmd_args.author, "subject": cmd_args.subject}
    else:
        if cmd_args.mode is None:
            main_parser.print_help()
            sys.exit(0)

    config_parameters = ["server", "username", "password", "primary_smtp_address", "ssl_verify", "target_path",
                         "author", "subject"]
    for parameter in config_parameters:
        if parameter not in config:
            logging.error("parameter " + parameter + " is not defined.")
            sys.exit(1)

    if not config["ssl_verify"]:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        exchangelib.BaseProtocol.HTTP_ADAPTER_CLS = exchangelib.NoVerifyHTTPAdapter

    exchange_credentials = exchangelib.Credentials(username=config["username"], password=config["password"])

    exchange_config = exchangelib.Configuration(server=config["server"], credentials=exchange_credentials)

    exchange_account = exchangelib.Account(primary_smtp_address=config["primary_smtp_address"], config=exchange_config,
                                           autodiscover=False, access_type=exchangelib.DELEGATE)

    exchange_account.public_folders_root.refresh()
    public_folders_root = exchange_account.public_folders_root
    target_folder = get_folder(config["target_path"], public_folders_root)

    item = exchangelib.PostItem(account=exchange_account, folder=target_folder, subject=config["subject"])
    item.author = config["author"]
    item.save()

    logging.info("Success")
    sys.exit(0)
