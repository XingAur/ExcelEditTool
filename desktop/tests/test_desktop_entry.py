import socket

from desktop.main import find_free_port


def test_find_free_port_returns_a_port_that_can_be_bound() -> None:
    port = find_free_port()

    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", port))
