#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
实现图片转表格功能

主要依赖腾讯云的OCR功能的行业文档-表格识别模块

使用时, 需指定环境变量`TC_SECRET_ID`和`TC_SECRET_KEY`, 否则报错

- Author: Rex Zhou <879582094@qq.com>
- Created Time: 2022/9/1 13:32
- Copyright: Copyright © 2022 Rex Zhou. All rights reserved.
"""

__version__ = "0.0.2"

__author__ = "Rex Zhou"
__copyright__ = "Copyright © 2022 Rex Zhou. All rights reserved."
__credits__ = [__author__]
__license__ = "None"
__maintainer__ = __author__
__email__ = "879582094@qq.com"

import base64
import json
import logging
import os
from pathlib import Path

import typer
from tencentcloud.common import credential
from tencentcloud.common.exception.tencent_cloud_sdk_exception import TencentCloudSDKException
from tencentcloud.common.profile.client_profile import ClientProfile
from tencentcloud.common.profile.http_profile import HttpProfile
from tencentcloud.ocr.v20181119 import ocr_client, models

ENCODING = 'utf-8'
# For more: https://cloud.tencent.com/document/product/866/49525
SUPPORT_SUFFIX = ['.png', '.jpg', '.jpeg', 'bmp', '.pdf']
MAX_SIZE = 7 * 1024**2

IMAGE_DIR = Path('tables')
OUTPUT = Path('output')


def get_client() -> ocr_client.OcrClient:
    """
    获取腾讯云认证客户端方法

    :return: ocr_client.OcrClient
    """
    # 实例化一个认证对象，入参需要传入腾讯云账户secretId，secretKey,此处还需注意密钥对的保密
    # 密钥可前往: https://console.cloud.tencent.com/cam/capi 网站进行获取
    secret_id = os.environ.get('TC_SECRET_ID')
    secret_key = os.environ.get('TC_SECRET_KEY')
    cred = credential.Credential(secret_id, secret_key)
    # 实例化一个http选项，可选的，没有特殊需求可以跳过
    http_profile = HttpProfile(endpoint='ocr.tencentcloudapi.com')
    # 实例化一个client选项，可选的，没有特殊需求可以跳过
    client_profile = ClientProfile(httpProfile=http_profile)
    # 实例化要请求产品的client对象,clientProfile是可选的
    client = ocr_client.OcrClient(cred, 'ap-beijing', client_profile)
    return client


def get_all_files() -> list[Path]:
    """
    获取所有支持的文件清单

    :return: list
    """
    result = []
    for root, _, files in os.walk(IMAGE_DIR):
        for file in files:
            file = Path(root) / file
            if file.suffix not in SUPPORT_SUFFIX:
                continue
            result.append(file)
    return result


def _load_file(file: Path) -> str:
    logger = logging.getLogger(_load_file.__name__)
    with open(file, 'rb') as file_:
        content = file_.read()
    content = base64.b64encode(content)
    if (size := len(content)) > MAX_SIZE:
        msg = f'base64加密后的文件大小超过限定值({MAX_SIZE/1024**2}MB): {size/1024**2}'
        logger.error(msg)
        return ''
    return str(content, encoding=ENCODING)


def get_ocr_result(client: ocr_client.OcrClient,
                   file: Path,
                   pdf_page: int = None) -> bytes:
    """
    获取OCR识别结果方法

    :param client: ocr_client.OcrClient 客户端实例
    :param file: Path 文件实例
    :param pdf_page: int 可选参数, 针对PDF文件指定页码, 默认为1
    :return:
    """
    logger = logging.getLogger(get_ocr_result.__name__)
    if not (content := _load_file(file)):
        return b''
    params = {'ImageBase64': content}
    if file.suffix == '.pdf':
        params['IsPdf'] = True
    if pdf_page:
        params['PdfPageNumber'] = pdf_page
    logger.debug('额外请求参数: %s',
                 {k: v for k, v in params.items() if k != 'ImageBase64'})
    # 实例化一个请求对象,每个接口都会对应一个request对象
    req = models.RecognizeTableOCRRequest()
    req.from_json_string(json.dumps(params))
    try:
        # 返回的resp是一个RecognizeTableOCRResponse的实例，与请求对象对应
        response = client.RecognizeTableOCR(req)
        data = base64.b64decode(response.Data)
        logger.debug('已获取结果, 请求ID: %s', response.RequestId)
        return data
    except TencentCloudSDKException as error:
        logger.error(error)
    return b''


def collect_data(client: ocr_client.OcrClient, files: list[Path]):
    """
    获取并汇总数据方法

    :param client: ocr_client.OcrClient 腾讯云OCR客户端实例
    :param files: list 文件列表
    :return: None
    """
    logger = logging.getLogger(collect_data.__name__)
    logger.info('开始获取并汇总OCR识别结果')
    OUTPUT.mkdir(exist_ok=True)
    for file in files:
        name = file.stem
        logger.info('正在处理: %s', name)
        result = get_ocr_result(client, file)
        excel = (OUTPUT / name).with_suffix('.xlsx')
        with open(excel, 'wb') as file_:
            file_.write(result)
        logger.info('已写入工作表: %s', name)


def main(debug: bool = False) -> None:
    """
    Main entry point function.

    :return:    :class:`None`
    """
    level = 'debug' if debug else 'info'
    log = Path(__file__).with_suffix('.log')
    logging.basicConfig(
        level=level.upper(),
        handlers=[logging.FileHandler(log),
                  logging.StreamHandler()],
        format='[%(asctime)s] [%(name)s] [%(levelname)s] %(message)s')
    logger = logging.getLogger(main.__name__)
    logger.info('正在获取文件清单')
    files = get_all_files()
    logger.info('已获取文件%d个', len(files))
    logger.debug([_.stem for _ in files])
    logger.info('正在获取腾讯OCR客户端')
    client = get_client()
    collect_data(client, files)


if __name__ == "__main__":
    typer.run(main)
