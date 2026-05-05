from rag_chatbot.main import main


def test_app_entrypoint_callable() -> None:
    assert callable(main)
