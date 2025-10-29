# SPROXY recipes

```python
from sapscriptwizard import SapSession, SProxy

session = SapSession.from_gui()
sproxy = SProxy(session)
sproxy.open_service("Z_MY_SERVICE")
```
