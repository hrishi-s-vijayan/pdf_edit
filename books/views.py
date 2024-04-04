from django.core.cache import cache

# Create your views here.

from rest_framework import viewsets

from .models import Book

from .serializers import BookSerializer


class BookViewSet(viewsets.ModelViewSet):
    queryset = Book.objects.all()
    serializer_class = BookSerializer

    def list(self, request, *args, **kwargs):
        cached_data = cache.get('book_list_data')
        if cached_data:
            return super().list(request, *args, **kwargs)
        

        result = calc()
        cache.set('book_list_data', result, timeout=10)

        return super().list(request, *args, **kwargs)
    

def calc():
    for i in Book.objects.all():
        # count = Book.objects.filter(name=i.name).count()
        print('===== count is : ', i)
    return 100
