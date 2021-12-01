import datetime
import hashlib
import json
import os
import random

from rest_framework import status, serializers
from rest_framework.generics import GenericAPIView, ListAPIView, get_object_or_404
from rest_framework.permissions import AllowAny, IsAuthenticated
from rest_framework.decorators import permission_classes, api_view
from rest_framework.response import Response
from rest_framework.views import APIView
from rest_framework.viewsets import ModelViewSet
from rest_framework import viewsets
from rest_framework_jwt.views import ObtainJSONWebToken
from django.core.paginator import Paginator 
from django.http import JsonResponse, HttpResponse 
import xlwt 
from vippu_backend.settings import unique_token

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

class BattallionOneViewset(viewsets.ModelViewSet):
    permission_classes = (IsAuthenticated,)
    serializer_class = BattallionOneSerializer
    queryset = Battallion_one.objects.all()

class BattalionTwo_overrall(GenericAPIView):
    permission_classes = (IsAuthenticated,)
    serializer_class = BattallionTwoSerializer

    def get(self, request):
        query_parameter_1 = "Embassy"
        query_parameter_2 = "Consolate"
        query_parameter_3 = "High commission"
        query_parameter_4 = "Other diplomats"
        query_parameter_5 = "Administration"

        embassy = len(Battallion_two.objects.filter(department=query_parameter_1))
        consolate = len(Battallion_two.objects.filter(department=query_parameter_2))
        high_commission = len(Battallion_two.objects.filter(department=query_parameter_3))
        other_diplomats = len(Battallion_two.objects.filter(department=query_parameter_4))
        administration = len(Battallion_two.objects.filter(department=query_parameter_5))

        summary_object = {
            "embassy": embassy,
            "consolate": consolate,
            "high_commission":  high_commission,
            "other_diplomats": other_diplomats,
            "administration": administration
        }
        return Response(summary_object, status=status.HTTP_200_OK)

class BattalionOne_overrall(GenericAPIView):
    permission_classes = (IsAuthenticated,)
    serializer_class = BattallionOneSerializer

    def get(self, request):
        query_parameter_1 = "UN Agencies"
        query_parameter_2 = "Administration"
        query_parameter_3 = "Drivers"

        agencies = len(Battallion_one.objects.filter(department=query_parameter_1))
        administration = len(Battallion_one.objects.filter(department=query_parameter_2))
        drivers = len(Battallion_one.objects.filter(department=query_parameter_3))
        

        summary_object = {
            "agencies": agencies,
            "administration": administration,
            "drivers":  drivers
            
        }
        return Response(summary_object, status=status.HTTP_200_OK)

class BattalionTwo_section_query(GenericAPIView):
    permission_classes = (IsAuthenticated,)
    serializer_class = BattallionOneSerializer

    def post(self, request):
        try:
            print(request.data['section'])
            query_parameter = request.data['section']
            query = Battallion_one.objects.filter(section=query_parameter)
            serializer = BattallionOneSerializer(query, many=True)
            return Response(serializer.data, status=status.HTTP_200_OK)
        except:
            content = {"error": "Section doesn't exit in the database"}
            return Response(content, status=status.HTTP_400_BAD_REQUEST)

class BattalionTwo_query(GenericAPIView):
    permission_classes = (IsAuthenticated,)
    serializer_class = BattallionTwoSerializer

    def post(self, request):
        print(request.data['file_number'])
        query_parameter = request.data['file_number']
        try:
            query = Battallion_two.objects.get(file_number=query_parameter)
            serializer = BattallionTwoSerializer(query, many=False)
            return Response(serializer.data, status=status.HTTP_200_OK)
        except:
            content = {"error": "Employee with this file number doesn't exit in the database"}
            return Response(content, status=status.HTTP_400_BAD_REQUEST)

class BattalionOne_query(GenericAPIView):
    permission_classes = (IsAuthenticated,)
    serializer_class = BattallionOneSerializer

    def post(self, request):
        print(request.data['file_number'])
        query_parameter = request.data['file_number']
        try:
            query = Battallion_one.objects.get(file_number=query_parameter)
            serializer = BattallionOneSerializer(query, many=False)
            return Response(serializer.data, status=status.HTTP_200_OK)
        except:
            content = {"error": "Employee with this file number doesn't exit in the database"}
            return Response(content, status=status.HTTP_400_BAD_REQUEST)

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

# @api_view(('GET',))      
def export_excel(request):
    try: 
        # to implement some security here 
        print(request.method)
        print(request.GET['unique'] )
        token = request.GET['unique']

        if token == unique_token:
            response = HttpResponse(content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename=Expenses' + \
                 str(datetime.datetime.now())+'.xls'
            wb = xlwt.Workbook(encoding='utf-8')
            ws = wb.add_sheet('Battalion Two')

            row_num = 1
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = [
                '', '', '', '','','','','',
                'BATTALION TWO DATA',
                '','','','','','','','','','','','','','','','','',
                '']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            row_num = 3
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = ['FILE NAME', 'LAST NAME', 'FILE NUMBER', 'NIN','IPPS',
                'ACCOUNT NUMBER',
                'TEL CONTACT',
                'SEX',
                'RANK',
                'EDUCATION LEVEL',
                'OTHER EDUCATION LEVEL',
                'BANK',
                'BRANCH',
                'DEPARTEMENT',
                'TITLE',
                'STATUS',
                'SHIFT',
                'DATE OF ENLISTMENT',
                'DATE OF TRANSFER',
                'DATE OF PROMOTION',
                'DATE OF BIRTH',
                'ARMED',
                'SECTION',
                'LOCATION',
                'ON LEAVE',
                'LEAVE START DATE',
                'LEAVE END DATE']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            font_style = xlwt.XFStyle()

            rows = Battallion_two.objects.filter().values_list('first_name', 'last_name', 'file_number', 'nin', 
                'ipps',
                'account_number',
                'contact',
                'sex',
                'rank',
                'education_level',
                'other_education_level',
                'bank',
                'branch',
                'department',
                'title',
                'status',
                'shift',
                'date_of_enlistment',
                'date_of_transfer',
                'date_of_promotion',
                'date_of_birth',
                'armed',
                'section',
                'location',
                'on_leave',
                'leave_start_date',
                'leave_end_date')

            for row in rows:
                row_num += 1

                for col_num in range(len(row)):
                    ws.write(row_num, col_num, str(row[col_num]), font_style)
            wb.save(response)

            return response
        else: 
            print("You are not authorized to carry out this operation.")
            return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)
    except: 
        print("You are not authorized to carry out this operation.")
        return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)

















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