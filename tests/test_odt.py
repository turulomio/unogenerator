from os import remove
from unogenerator import ODT_Standard, ODT


def test_metadata():
    with ODT_Standard() as doc:
        doc.setMetadata(
            title="Metadata example",
            subject="UnoGenerator testing",
            author="Turulomio",
            description="This testing works with metadata methods",
            keywords=["Metadata", "Testing"],
        )
        doc.save("delete_metadata.odt")

    with ODT("delete_metadata.odt") as doc:
        assert doc.getMetadata()["Title"]== "Metadata example"
        
    remove("delete_metadata.odt")
