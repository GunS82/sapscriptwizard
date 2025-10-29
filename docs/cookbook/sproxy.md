# SPROXY navigation

```python
from sapscriptwizard import SapSession
from sapscriptwizard.features.sproxy import SPROXY

session = SapSession()
session.connect()
window = session.window()
SPROXY(window).navigate_repository(
    "wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell", ["Root", "Namespace", "Service"]
)
```
