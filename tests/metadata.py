from datetime import datetime
from unogenerator import ODT_Standard, ODT

with ODT_Standard() as doc:
    doc.setMetadata(
        title="Metadata example",
        subject="Unogenerator testing",
        author="Turulomio",
        description="This testing works with metadata methods",
        keywords=["Metadata", "Testing"],
    )
    doc.save("delete_metadata.odt")

with ODT("delete_metadata.odt") as doc:
    print(doc.getMetadata())
