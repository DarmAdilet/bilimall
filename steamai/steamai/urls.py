from django.contrib import admin
from django.urls import path,include
from django.conf import settings
from django.conf.urls.static import static  # осыны қосу маңызды!
urlpatterns = [
    path('admin/', admin.site.urls),
    path('',include('steam.urls'))
]
if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
