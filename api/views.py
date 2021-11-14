import datetime
import hashlib
import json
import os
import random

from rest_framework import status, serializers
from rest_framework.generics import GenericAPIView, ListAPIView, get_object_or_404
from rest_framework.permissions import AllowAny, IsAuthenticated
from rest_framework.response import Response
from rest_framework.views import APIView
from rest_framework.viewsets import ModelViewSet
from rest_framework import viewsets
from rest_framework_jwt.views import ObtainJSONWebToken
from django.core.paginator import Paginator 
from django.http import JsonResponse, HttpResponse 
import xlwt 

from api.models import *
from api.serializers import *


class AccountLoginAPIView(ObtainJSONWebToken):
    serializer_class = JWTSerializer

class UserType(GenericAPIView):
    permission_classes = (AllowAny,)
    serializer_class = UserSerializer

    def post(self, request):
        username = request.data["username"]
        # print(username)
        try: 
            user = User.objects.get(username=username)
            user_data = UserSerializer(user, many=False).data
            # print(user_data)
            user_type = user_data["account_type"]
            if(user_type == 'admin'):
                return Response({
                    "user_type": "admin",
                    })
            elif(user_type == 'in_charge'):
                return Response({
                    "user_type": "in_charge",
                    })
            else: 
                return Response({
                    "user_type": "none",
                    })
        except: 
            print("Not found")
            return Response({
                    "user_type": "none",
                    })

class SignUp(GenericAPIView):
    permission_classes = (AllowAny,)
    serializer_class = SignUpSerializer

    def post(self, request):
        serializer = self.serializer_class(data=request.data)
        if serializer.is_valid():
            user = serializer.save()
            user.is_active = False
            user.set_password(serializer.validated_data["password"])
            user.save()

            username = serializer.data["username"]

            # Get user by username
            user = User.objects.get(username=username)

            return Response(
                UserSerializer(user, many=False).data, status=status.HTTP_200_OK
            )
        else:
            return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)

class UserProfile(GenericAPIView):
    permission_classes = (IsAuthenticated,)
    serializer_class = UserSerializer

    def get(self, request):
        user = request.user
        serializer = UserSerializer(user, many=False)
        return Response(serializer.data, status=status.HTTP_200_OK)

class BattallionTwoViewset(viewsets.ModelViewSet):
    permission_classes = (IsAuthenticated,)
    serializer_class = BattallionTwoSerializer
    queryset = Battallion_two.objects.all()

    

class ChangePasswordApi(GenericAPIView):
    serializer_class = ChangePasswordSerializer

    def post(self, request):
        serializer = self.serializer_class(data=request.data)

        if serializer.is_valid():
            old_password = serializer.data["old_password"]
            new_password = serializer.data["new_password"]

            user = self.request.user

            if not user.check_password(old_password):
                content = {"detail": "Invalid Password"}
                return Response(content, status=status.HTTP_400_BAD_REQUEST)
            else:
                user.set_password(new_password)
                user.save()
                content = {"success": "Password Changed"}
                return Response(content, status=status.HTTP_200_OK)
        else:
            return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)

def export_excel(request):

    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename=Expenses' + \
         str(datetime.datetime.now())+'.xls'
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('Battalion Two')

    row_num = 0
    font_style = xlwt.XFStyle()
    font_style.font.bold = True

    columns = ['First name', 'Last name', 'File number', 'Nin']

    for col_num in range(len(columns)):
        ws.write(row_num, col_num, columns[col_num], font_style)

    font_style = xlwt.XFStyle()

    rows = Battallion_two.objects.filter().values_list('first_name', 'last_name', 'file_number', 'nin')

    for row in rows:
        row_num += 1

        for col_num in range(len(row)):
            ws.write(row_num, col_num, str(row[col_num]), font_style)
    wb.save(response)

    return response

















# def get_permissions(self):
#     if self.action == "retrieve":
#         self.permission_classes = [AllowAny]
#     else:
#         self.permission_classes = [IsAuthenticated]

#     return super(UserViewSet, self).get_permissions()

# def list(self, request):
#     return Response(status=status.HTTP_405_METHOD_NOT_ALLOWED)

# def perform_create(self, serializer):
#     user = serializer.save()

#     customer_data = stripe.Customer.list(email=user.email).data

#     # if the array is empty it means the email has not been used yet
#     if len(customer_data) == 0:
#         # creating stripe customer
#         customer = stripe.Customer.create(email=user.email, payment_method="card")

# def partial_update(self, request, *args, **kwargs):
#     instance = self.queryset.get(pk=kwargs.get('pk'))
#     serializer = self.serializer_class(instance, data=request.data, partial=True)
#     serializer.is_valid(raise_exception=True)

#     date_of_birth = serializer.validated_data.get('date_of_birth', None)
#     if date_of_birth:
#         if date_of_birth >= datetime.datetime.today().date():
#             raise serializers.ValidationError(
#                 {"date_of_birth": "Date of Birth cannot be greater than or equal to current date"}
#             )
#     serializer.save()
#     return Response(serializer.data)
# def destroy(self, request, *args, **kwargs):
#     payment_method = self.get_object()

#     try:
#         stripe.PaymentMethod.detach(payment_method.stripe_id)
#         payment_method.delete()
#     except Exception as e:
#         print(e)
#     return Response(status=status.HTTP_204_NO_CONTENT)