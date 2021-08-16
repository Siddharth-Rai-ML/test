import json

from requests import get, post

from logger import log
import constants


class API:
    def __init__(self):
        self.token = self.__get_token()

    def __get_token(self):
        data = {'username': constants.NEW_TWISTLOCK_USERNAME,
                'password': constants.NEW_TWISTLOCK_PASSWORD}
        log.debug('Fetching api token...')
        resp = self.__api_post_auth_request(data)
        try:
            return resp['token']
        except KeyError:
            raise Exception('KeyError. Auth Response does not have a token')
        except Exception as e:
            exception_type = type(e).__name__
            exception_args = e.args
            raise Exception(
                f'UNKOWN ERROR - __get_token - ExceptionType: {exception_type}. Exception Args: {exception_args}')

    def __api_post_auth_request(self, params: dict):
        endpoint = constants.AUTH_ENDPOINT
        log.debug('making post request...')
        return self.__api_post_request({}, params, endpoint)

    def __api_post_request(self, headers: dict, params: dict, endpoint: str):
        url = constants.NEW_TWISTLOCK_BASE_URL + "/" + endpoint
        base_headers = {'Content-Type': 'application/json'}
        cust_headers = {**base_headers, **headers}
        log.debug(f'posting to {url}')
        resp = post(
            url,
            headers=cust_headers,
            data=json.dumps(params),
            verify=False
        )
        # print(resp, resp.text)
        if resp.status_code == 401:
            raise Exception('Invalid credentials')
        elif resp.status_code == 500:
            raise Exception('Internal Server Error')
        elif resp.status_code == 505:
            raise Exception('Server Error: Service Unavailable')
        try:
            return resp.json()
        except ValueError:
            raise Exception('Unable to JSON decode API POST response')
        except Exception as e:
            exception_type = type(e).__name__
            exception_args = e.args
            raise Exception(
                f'UNKOWN ERROR. __api_post_request - ExceptionType: {exception_type}. Exception Args: {exception_args}')

    def __api_get_request(self, headers: dict, endpoint: str):
        base_headers = {'Content-Type': 'application/json'}
        cust_headers = {**base_headers, **headers}
        url = constants.NEW_TWISTLOCK_BASE_URL + "/" + endpoint
        log.debug(f'getting from {url}')
        resp = get(
            url,
            headers=cust_headers,
            verify=False
        )
        # print(resp, resp.text, type(resp.text))
        if resp.status_code == 401:
            raise Exception('Invalid credentials')
        elif resp.status_code == 500:
            raise Exception('Internal Server Error')
        elif resp.status_code == 505:
            raise Exception('Server Error: Service Unavailable')
        try:
            return resp.json()
        except ValueError:
            raise Exception('Unable to JSON decode API POST response')
        except Exception as e:
            exception_type = type(e).__name__
            exception_args = e.args
            raise Exception(
                f'UNKOWN ERROR. __api_post_request - ExceptionType: {exception_type}. Exception Args: {exception_args}')

    def __api_get_offset_request(self, headers: dict, endpoint: str):
        base_headers = {'Content-Type': 'application/json'}
        cust_headers = {**base_headers, **headers}
        url = constants.NEW_TWISTLOCK_BASE_URL + "/" + endpoint

        log.debug(f'offset request base url - {url}')

        resp_list = []
        resp_list.sort()
        offset = 0

        ###########################################
        #while True:
        for i in range(1):
        ###########################################
            response = get(url + "?limit=50&offset={}".format(offset), headers=cust_headers, verify=False)
            offset += 50
            log.debug(f"offset - {offset}")
            if response.json() is None:
                log.debug("Got null response existing the while loop")
                break
            else:
                log.debug("Still in while loop added response output to list")
                resp_list += response.json()
        return resp_list

    def get_registries(self):
        endpoint = constants.REGISTRY_ENDPOINT
        token = self.__get_token()
        headers = {'Authorization': f'Bearer {token}'}
        log.debug("getting registries...")
        return self.__api_get_offset_request(headers, endpoint)

    def get_collection(self):
        endpoint = constants.COLLECTION_ENDPOINT
        token = self.__get_token()
        headers = {'Authorization': f'Bearer {token}'}
        log.debug("getting collection...")
        return self.__api_get_request(headers, endpoint)


api = API()
