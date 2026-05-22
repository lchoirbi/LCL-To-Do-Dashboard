from __future__ import annotations

import time

import requests


class PublicClientApplication:
    def __init__(self, client_id: str, authority: str) -> None:
        self.client_id = client_id
        self.authority = authority.rstrip("/")

    def initiate_device_flow(self, scopes: list[str]):
        response = requests.post(
            f"{self.authority}/oauth2/v2.0/devicecode",
            data={"client_id": self.client_id, "scope": " ".join(scopes)},
            timeout=30,
        )
        return response.json()

    def acquire_token_by_device_flow(self, flow: dict):
        response = requests.post(
            f"{self.authority}/oauth2/v2.0/token",
            data={
                "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
                "client_id": self.client_id,
                "device_code": flow.get("device_code"),
            },
            timeout=30,
        )
        result = response.json()
        if "access_token" in result:
            result["expires_on"] = int(time.time()) + int(result.get("expires_in", 3600))
        return result
