from __future__ import annotations

import httpx
import pytest
import respx

from oikb.cli import _resolve_connector
from oikb.connectors.sharepoint import SharePointConnector

GRAPH = "https://graph.microsoft.com/v1.0"
SOURCE = "sharepoint:contoso.sharepoint.com/sites/Team/Docs"


@pytest.fixture(autouse=True)
def _sharepoint_env(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("SHAREPOINT_TENANT_ID", "tenant")
    monkeypatch.setenv("SHAREPOINT_CLIENT_ID", "client")
    monkeypatch.setenv("SHAREPOINT_CLIENT_SECRET", "secret")


def _mock_graph(router: respx.MockRouter) -> None:
    router.post("https://login.microsoftonline.com/tenant/oauth2/v2.0/token").respond(
        json={"access_token": "token"}
    )
    router.get(f"{GRAPH}/sites/contoso.sharepoint.com:/sites/Team").respond(json={"id": "site-id"})
    router.get(f"{GRAPH}/sites/site-id/drives").respond(json={"value": [{"id": "drive-id", "name": "Docs"}]})
    router.get(f"{GRAPH}/drives/drive-id/root/children").respond(
        json={
            "value": [
                {"name": "Exceptions", "folder": {}},
                {"name": "top.pdf", "file": {}, "eTag": '"top-etag"', "size": 1},
            ]
        }
    )
    router.get(f"{GRAPH}/drives/drive-id/root:/Exceptions:/children").respond(
        json={
            "value": [
                {"name": "VM_POWERED_OFF.xlsx", "file": {}, "eTag": '"vm-etag"', "size": 2},
                {"name": "nested", "folder": {}},
            ]
        }
    )
    router.get(f"{GRAPH}/drives/drive-id/root:/Exceptions/nested:/children").respond(
        json={"value": [{"name": "deep.xlsx", "file": {}, "eTag": '"deep-etag"', "size": 3}]}
    )


@respx.mock(assert_all_called=False)
def test_default_root_is_library_root(respx_mock: respx.MockRouter) -> None:
    _mock_graph(respx_mock)

    manifest = _resolve_connector(SOURCE).build_manifest()

    assert [e.display_path for e in manifest] == [
        "Exceptions/VM_POWERED_OFF.xlsx",
        "Exceptions/nested/deep.xlsx",
        "top.pdf",
    ]


@respx.mock(assert_all_called=False)
def test_root_path_scopes_manifest_and_strips_the_folder(respx_mock: respx.MockRouter) -> None:
    _mock_graph(respx_mock)

    manifest = _resolve_connector(SOURCE, path="Exceptions").build_manifest()

    assert [e.display_path for e in manifest] == [
        "VM_POWERED_OFF.xlsx",
        "nested/deep.xlsx",
    ]


@respx.mock(assert_all_called=False)
def test_root_path_tolerates_surrounding_slashes(respx_mock: respx.MockRouter) -> None:
    _mock_graph(respx_mock)

    manifest = _resolve_connector(SOURCE, path="/Exceptions/").build_manifest()

    assert [e.display_path for e in manifest] == ["VM_POWERED_OFF.xlsx", "nested/deep.xlsx"]


@respx.mock(assert_all_called=False)
def test_read_file_resolves_paths_against_root_path(respx_mock: respx.MockRouter) -> None:
    _mock_graph(respx_mock)
    root_file = respx_mock.get(f"{GRAPH}/drives/drive-id/root:/Exceptions/VM_POWERED_OFF.xlsx:/content").respond(
        content=b"top"
    )
    nested_file = respx_mock.get(f"{GRAPH}/drives/drive-id/root:/Exceptions/nested/deep.xlsx:/content").respond(
        content=b"deep"
    )
    connector = SharePointConnector(site="contoso.sharepoint.com", site_path="sites/Team", library="Docs", root_path="Exceptions")

    assert connector.read_file("", "VM_POWERED_OFF.xlsx") == b"top"
    assert connector.read_file("nested", "deep.xlsx") == b"deep"
    assert root_file.called and nested_file.called


@respx.mock(assert_all_called=False)
def test_missing_root_path_raises(respx_mock: respx.MockRouter) -> None:
    _mock_graph(respx_mock)
    respx_mock.get(f"{GRAPH}/drives/drive-id/root:/Nope:/children").respond(
        404, json={"error": {"code": "itemNotFound"}}
    )

    with pytest.raises(httpx.HTTPStatusError):
        _resolve_connector(SOURCE, path="Nope").build_manifest()
