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
            commander = user_data["top_level_incharge"]
            if(user_type == 'admin'):
                return Response({
                    "user_type": "admin",
                    })
            elif(user_type == 'in_charge'):
                if(commander == True):
                    return Response({
                        "user_type": "commander",
                        })
                else:
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

class BattallionThreeViewset(viewsets.ModelViewSet):
    permission_classes = (IsAuthenticated,)
    serializer_class = BattallionThreeSerializer
    queryset = Battallion_three.objects.all()

class BattallionFourViewset(viewsets.ModelViewSet):
    permission_classes = (IsAuthenticated,)
    serializer_class = BattallionFourSerializer
    queryset = Battallion_four.objects.all()

class BattallionFiveViewset(viewsets.ModelViewSet):
    permission_classes = (IsAuthenticated,)
    serializer_class = BattallionFiveSerializer
    queryset = Battallion_five.objects.all()

class DeletedEmployeeViewset(viewsets.ModelViewSet):
    permission_classes = (IsAuthenticated,)
    serializer_class = DeletedEmployeeSerializer
    queryset = Deleted_Employee.objects.all()

class BattallionOneViewset(viewsets.ModelViewSet):
    permission_classes = (IsAuthenticated,)
    serializer_class = BattallionOneSerializer
    queryset = Battallion_one.objects.all()


class BattalionThree_overrall(GenericAPIView):
    permission_classes = (IsAuthenticated,)
    serializer_class = BattallionThreeSerializer

    def get(self, request):
        query_parameter_1 = "Anti-corruption and War Crime division"
        query_parameter_2 = "Buganda Road Court"
        query_parameter_3 = "Commercial Court"
        query_parameter_4 = "Supreme Court"
        query_parameter_5 = "High Court Offices Kampala"
        query_parameter_6 = "High Court Residence"
        query_parameter_7 = "Family Court Division Makindye"
        query_parameter_8 = "Court of Appeal"
        query_parameter_9 = "Residence of Justice of Court Appeal"
        query_parameter_10 = "DPP Office"
        query_parameter_11 = "IGG"
        query_parameter_12 = "Police Establishment"
        query_parameter_13 = "AOG"

        total = len(Battallion_three.objects.all())
        corruption = len(Battallion_three.objects.filter(department=query_parameter_1))
        buganda_road_court = len(Battallion_three.objects.filter(department=query_parameter_2))
        commercial_court = len(Battallion_three.objects.filter(department=query_parameter_3))
        supreme_court = len(Battallion_three.objects.filter(department=query_parameter_4))
        high_court_offices = len(Battallion_three.objects.filter(department=query_parameter_5))
        high_court_residence = len(Battallion_three.objects.filter(department=query_parameter_6))
        family_court_division = len(Battallion_three.objects.filter(department=query_parameter_7))
        court_of_appeal = len(Battallion_three.objects.filter(department=query_parameter_8))
        residence_of_justice = len(Battallion_three.objects.filter(department=query_parameter_9))
        dpp_office = len(Battallion_three.objects.filter(department=query_parameter_10))
        igg = len(Battallion_three.objects.filter(department=query_parameter_11))
        police_establishment = len(Battallion_three.objects.filter(department=query_parameter_12))
        aog = len(Battallion_three.objects.filter(department=query_parameter_13))

        summary_object = {
            "total": total,
            "corruption": corruption,
            "buganda_road_court": buganda_road_court,
            "commercial_court":  commercial_court,
            "supreme_court": supreme_court,
            "high_court_offices": high_court_offices,
            "high_court_residence": high_court_residence,
            "family_court_division": family_court_division,
            "court_of_appeal": court_of_appeal,
            "residence_of_justice": residence_of_justice,
            "dpp_office": dpp_office,
            "igg": igg,
            "aog": aog,
            "police_establishment": police_establishment
        }
        return Response(summary_object, status=status.HTTP_200_OK)

class BattalionFive_overrall(GenericAPIView):
    permission_classes = (IsAuthenticated,)
    serializer_class = BattallionFiveSerializer

    def get(self, request):
        query_parameter_1 = "UCC"
        query_parameter_2 = "EC"
        query_parameter_3 = "IRA"
        query_parameter_4 = "URA"
        query_parameter_5 = "UNRA"
        query_parameter_6 = "NPA"
        query_parameter_7 = "ULC"
        query_parameter_8 = "PSC"
        query_parameter_9 = "NSSF"
        query_parameter_10 = "KCCA"
        query_parameter_11 = "SENIOR CITIZENS"
        query_parameter_12 = "JSC"
        query_parameter_13 = "EOC"
        query_parameter_14 = "Administration"

        total = len(Battallion_five.objects.all())
        ucc = len(Battallion_five.objects.filter(department=query_parameter_1))
        ec = len(Battallion_five.objects.filter(department=query_parameter_2))
        ira = len(Battallion_five.objects.filter(department=query_parameter_3))
        ura = len(Battallion_five.objects.filter(department=query_parameter_4))
        unra = len(Battallion_five.objects.filter(department=query_parameter_5))
        npa = len(Battallion_five.objects.filter(department=query_parameter_6))
        ulc = len(Battallion_five.objects.filter(department=query_parameter_7))
        psc = len(Battallion_five.objects.filter(department=query_parameter_8))
        nssf = len(Battallion_five.objects.filter(department=query_parameter_9))
        kcca = len(Battallion_five.objects.filter(department=query_parameter_10))
        senior_citizens = len(Battallion_five.objects.filter(department=query_parameter_11))
        jsc = len(Battallion_five.objects.filter(department=query_parameter_12))
        eoc = len(Battallion_five.objects.filter(department=query_parameter_13))
        administration = len(Battallion_five.objects.filter(department=query_parameter_14))

        summary_object = {
            "total": total,
            "ucc": ucc,
            "ec": ec,
            "ira": ira,
            "ura": ura,
            "unra": unra,
            "npa": npa,
            "ulc": ulc,
            "psc": psc,
            "nssf": nssf,
            "kcca": kcca,
            "senior_citizens": senior_citizens,
            "jsc": jsc,
            "eoc": eoc,
            "administration": administration
        }
        return Response(summary_object, status=status.HTTP_200_OK)

class BattalionFour_overrall(GenericAPIView):
    permission_classes = (IsAuthenticated,)
    serializer_class = BattallionFourSerializer

    def get(self, request):
        query_parameter_1 = "Body guard"
        query_parameter_2 = "Crew Commander"
        query_parameter_3 = "Crew"

        total = len(Battallion_four.objects.all())
        body_gaurd = len(Battallion_four.objects.filter(department=query_parameter_1))
        crew_commander = len(Battallion_four.objects.filter(department=query_parameter_2))
        crew = len(Battallion_four.objects.filter(department=query_parameter_3))

        summary_object = {
            "total": total,
            "body_gaurd": body_gaurd,
            "crew_commander": crew_commander,
            "crew":  crew
        }
        return Response(summary_object, status=status.HTTP_200_OK)

class BattalionTwo_overrall(GenericAPIView):
    permission_classes = (IsAuthenticated,)
    serializer_class = BattallionTwoSerializer

    def get(self, request):
        query_parameter_1 = "Embassy"
        query_parameter_2 = "Consulate"
        query_parameter_3 = "High Commission"
        query_parameter_4 = "Other Diplomats"
        query_parameter_5 = "Administration"

        total = len(Battallion_two.objects.all())
        embassy = len(Battallion_two.objects.filter(department=query_parameter_1))
        consolate = len(Battallion_two.objects.filter(department=query_parameter_2))
        high_commission = len(Battallion_two.objects.filter(department=query_parameter_3))
        other_diplomats = len(Battallion_two.objects.filter(department=query_parameter_4))
        administration = len(Battallion_two.objects.filter(department=query_parameter_5))

        summary_object = {
            "total": total,
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

        total = len(Battallion_one.objects.all())
        agencies = len(Battallion_one.objects.filter(department=query_parameter_1))
        administration = len(Battallion_one.objects.filter(department=query_parameter_2))
        drivers = len(Battallion_one.objects.filter(department=query_parameter_3))
        

        summary_object = {
            "total": total,
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
            # print(request.data['section'])
            query_parameter = request.data['section']
            query = Battallion_one.objects.filter(section=query_parameter)
            serializer = BattallionOneSerializer(query, many=True)
            return Response(serializer.data, status=status.HTTP_200_OK)
        except:
            content = {"error": "Section doesn't exit in the database"}
            return Response(content, status=status.HTTP_400_BAD_REQUEST)

class BattalionThree_department_query(GenericAPIView):
    permission_classes = (IsAuthenticated,)
    serializer_class = BattallionThreeSerializer

    def post(self, request):
        try:
            # print(request.data['department'])
            query_parameter = request.data['department']
            query = Battallion_three.objects.filter(department=query_parameter)
            serializer = BattallionThreeSerializer(query, many=True)
            return Response(serializer.data, status=status.HTTP_200_OK)
        except:
            content = {"error": "Department doesn't exit in the database"}
            return Response(content, status=status.HTTP_400_BAD_REQUEST)

class BattalionTwo_department_query(GenericAPIView):
    permission_classes = (IsAuthenticated,)
    serializer_class = BattallionTwoSerializer

    def post(self, request):
        try:
            # print(request.data['department'])
            query_parameter = request.data['department']
            query = Battallion_two.objects.filter(department=query_parameter)
            serializer = BattallionTwoSerializer(query, many=True)
            return Response(serializer.data, status=status.HTTP_200_OK)
        except:
            content = {"error": "Department doesn't exit in the database"}
            return Response(content, status=status.HTTP_400_BAD_REQUEST)

class BattalionFive_department_query(GenericAPIView):
    permission_classes = (IsAuthenticated,)
    serializer_class = BattallionFiveSerializer

    def post(self, request):
        try:
            # print(request.data['department'])
            query_parameter = request.data['department']
            query = Battallion_five.objects.filter(department=query_parameter)
            serializer = BattallionFiveSerializer(query, many=True)
            return Response(serializer.data, status=status.HTTP_200_OK)
        except:
            content = {"error": "Department doesn't exit in the database"}
            return Response(content, status=status.HTTP_400_BAD_REQUEST)

class BattalionFour_department_query(GenericAPIView):
    permission_classes = (IsAuthenticated,)
    serializer_class = BattallionFourSerializer

    def post(self, request):
        try:
            # print(request.data['department'])
            query_parameter = request.data['department']
            query = Battallion_four.objects.filter(department=query_parameter)
            serializer = BattallionFourSerializer(query, many=True)
            return Response(serializer.data, status=status.HTTP_200_OK)
        except:
            content = {"error": "Department doesn't exit in the database"}
            return Response(content, status=status.HTTP_400_BAD_REQUEST)

class BattalionTwo_query(GenericAPIView):
    permission_classes = (IsAuthenticated,)
    serializer_class = BattallionTwoSerializer

    def post(self, request):
        # print(request.data['file_number'])
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
        # print(request.data['file_number'])
        query_parameter = request.data['file_number']
        try:
            query = Battallion_one.objects.get(file_number=query_parameter)
            serializer = BattallionOneSerializer(query, many=False)
            return Response(serializer.data, status=status.HTTP_200_OK)
        except:
            content = {"error": "Employee with this file number doesn't exit in the database"}
            return Response(content, status=status.HTTP_400_BAD_REQUEST)

class BattalionThree_query(GenericAPIView):
    permission_classes = (IsAuthenticated,)
    serializer_class = BattallionThreeSerializer

    def post(self, request):
        # print(request.data['file_number'])
        query_parameter = request.data['file_number']
        try:
            query = Battallion_three.objects.get(file_number=query_parameter)
            serializer = BattallionThreeSerializer(query, many=False)
            return Response(serializer.data, status=status.HTTP_200_OK)
        except:
            content = {"error": "Employee with this file number doesn't exit in the database"}
            return Response(content, status=status.HTTP_400_BAD_REQUEST)


class BattalionFive_query(GenericAPIView):
    permission_classes = (IsAuthenticated,)
    serializer_class = BattallionFiveSerializer

    def post(self, request):
        # print(request.data['file_number'])
        query_parameter = request.data['file_number']
        try:
            query = Battallion_five.objects.get(file_number=query_parameter)
            serializer = BattallionFiveSerializer(query, many=False)
            return Response(serializer.data, status=status.HTTP_200_OK)
        except:
            content = {"error": "Employee with this file number doesn't exit in the database"}
            return Response(content, status=status.HTTP_400_BAD_REQUEST)

class BattalionFour_query(GenericAPIView):
    permission_classes = (IsAuthenticated,)
    serializer_class = BattallionFourSerializer

    def post(self, request):
        # print(request.data['file_number'])
        query_parameter = request.data['file_number']
        try:
            query = Battallion_four.objects.get(file_number=query_parameter)
            serializer = BattallionFourSerializer(query, many=False)
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


    
def export_excel(request):
    try: 
        # to implement some security here 
        # print(request.method)
        # print(request.GET['unique'] )
        print(request.user)
        token = request.GET['unique']
        title_doc = request.GET['title_doc']

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
                title_doc,
                '','','','','','','','','','','','','','','','','',
                '']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            row_num = 3
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = ['FILE NAME', 'LAST NAME', 'FILE NUMBER', 'NIN','IPPS','ACCOUNT NUMBER','TEL CONTACT','SEX','RANK','EDUCATION LEVEL','OTHER EDUCATION LEVEL',
            'BANK','BRANCH','DEPARTEMENT','TITLE','STATUS','SHIFT','DATE OF ENLISTMENT','DATE OF TRANSFER','DATE OF PROMOTION','DATE OF BIRTH','ARMED','SECTION','LOCATION','ON LEAVE','LEAVE START DATE',
                'LEAVE END DATE', 'SPECIAL DUTY START DATE', 'SPECIAL DUTY END DATE']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            font_style = xlwt.XFStyle()
            
            # for reports with specific Sections
            # query_parameter = "UN Women"
            # rows = Battallion_one.objects.filter(section=query_parameter).values_list('first_name', 'last_name', 'file_number', 'nin', 

            rows = Battallion_two.objects.filter().values_list('first_name', 'last_name', 'file_number', 'nin', 'ipps','account_number','contact','sex','rank','education_level','other_education_level','bank','branch',
                'department','title','status','shift','date_of_enlistment','date_of_transfer','date_of_promotion','date_of_birth','armed','section','location','on_leave','leave_start_date',
                'leave_end_date', 'special_duty_start_date', 'special_duty_end_date')

            for row in rows:
                row_num += 1

                for col_num in range(len(row)):
                    ws.write(row_num, col_num, str(row[col_num]), font_style)
            wb.save(response)

            return response
        else: 
            # print("You are not authorized to carry out this operation.")
            return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)
    except: 
        # print("You are not authorized to carry out this operation.")
        return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)


def export_excel_B_five_department_leave(request):
    try: 
        # to implement some security here 
        # print(request.method)
        # print(request.GET['leave_type'] )
        # print(request.GET['department'] )
        token = request.GET['unique']
        query_parameter = request.GET['department']
        query_parameter2 = request.GET['leave_type']
        title_doc = request.GET['title_doc']

        # title = 'BATTALION ONE ' + query_parameter + ' DATA'
        # print(title)

        if token == unique_token:
            response = HttpResponse(content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename=Expenses' + \
                 str(datetime.datetime.now())+'.xls'
            wb = xlwt.Workbook(encoding='utf-8')
            ws = wb.add_sheet('Battalion Five')

            row_num = 1
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = [
                '', '', '', '','','','','',
                title_doc,
                '','','','','','','','','','','','','','','','','',
                '']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            row_num = 3
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = ['FILE NAME', 'LAST NAME', 'FILE NUMBER', 'NIN','IPPS',
                'ACCOUNT NUMBER','TEL CONTACT','SEX','RANK','EDUCATION LEVEL','OTHER EDUCATION LEVEL','BANK','BRANCH','DEPARTEMENT','TITLE','STATUS','SHIFT','DATE OF ENLISTMENT','DATE OF TRANSFER','DATE OF PROMOTION','DATE OF BIRTH','ARMED','SECTION','LOCATION','ON LEAVE','LEAVE START DATE',
                'LEAVE END DATE', 'SPECIAL DUTY START DATE', 'SPECIAL DUTY END DATE']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            font_style = xlwt.XFStyle()
            
            rows = Battallion_five.objects.filter(department=query_parameter, on_leave=query_parameter2).values_list('first_name', 'last_name', 'file_number', 'nin', 
                'ipps','account_number','contact','sex','rank','education_level','other_education_level','bank','branch','department','title','status','shift','date_of_enlistment','date_of_transfer','date_of_promotion','date_of_birth','armed','section','location','on_leave','leave_start_date',
                'leave_end_date', 'special_duty_start_date', 'special_duty_end_date')

            for row in rows:
                row_num += 1

                for col_num in range(len(row)):
                    ws.write(row_num, col_num, str(row[col_num]), font_style)
            wb.save(response)

            return response
        else: 
            # print("You are not authorized to carry out this operation.")
            return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)
    except: 
        # print("You are not authorized to carry out this operation.")
        return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)


def export_excel_B_four_department_leave(request):
    try: 
        # to implement some security here 
        # print(request.method)
        # print(request.GET['leave_type'] )
        # print(request.GET['department'] )
        token = request.GET['unique']
        query_parameter = request.GET['department']
        query_parameter2 = request.GET['leave_type']
        title_doc = request.GET['title_doc']

        # title = 'BATTALION ONE ' + query_parameter + ' DATA'
        # print(title)

        if token == unique_token:
            response = HttpResponse(content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename=Expenses' + \
                 str(datetime.datetime.now())+'.xls'
            wb = xlwt.Workbook(encoding='utf-8')
            ws = wb.add_sheet('Battalion Four')

            row_num = 1
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = [
                '', '', '', '','','','','',
                title_doc,
                '','','','','','','','','','','','','','','','','',
                '']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            row_num = 3
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = ['FILE NAME', 'LAST NAME', 'FILE NUMBER', 'NIN','IPPS',
                'ACCOUNT NUMBER','TEL CONTACT','SEX','RANK','EDUCATION LEVEL','OTHER EDUCATION LEVEL','BANK','BRANCH','DEPARTEMENT','TITLE','STATUS','SHIFT','DATE OF ENLISTMENT','DATE OF TRANSFER','DATE OF PROMOTION','DATE OF BIRTH','ARMED','SECTION','LOCATION','ON LEAVE','LEAVE START DATE',
                'LEAVE END DATE', 'SPECIAL DUTY START DATE', 'SPECIAL DUTY END DATE']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            font_style = xlwt.XFStyle()
            
            rows = Battallion_four.objects.filter(department=query_parameter, on_leave=query_parameter2).values_list('first_name', 'last_name', 'file_number', 'nin', 
                'ipps','account_number','contact','sex','rank','education_level','other_education_level','bank','branch','department','title','status','shift','date_of_enlistment','date_of_transfer','date_of_promotion','date_of_birth','armed','section','location','on_leave','leave_start_date',
                'leave_end_date', 'special_duty_start_date', 'special_duty_end_date')

            for row in rows:
                row_num += 1

                for col_num in range(len(row)):
                    ws.write(row_num, col_num, str(row[col_num]), font_style)
            wb.save(response)

            return response
        else: 
            # print("You are not authorized to carry out this operation.")
            return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)
    except: 
        # print("You are not authorized to carry out this operation.")
        return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)


def export_excel_B_three_department_leave(request):
    try: 
        # to implement some security here 
        # print(request.method)
        # print(request.GET['leave_type'] )
        # print(request.GET['department'] )
        token = request.GET['unique']
        query_parameter = request.GET['department']
        query_parameter2 = request.GET['leave_type']
        title_doc = request.GET['title_doc']

        # title = 'BATTALION ONE ' + query_parameter + ' DATA'
        # print(title)

        if token == unique_token:
            response = HttpResponse(content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename=Expenses' + \
                 str(datetime.datetime.now())+'.xls'
            wb = xlwt.Workbook(encoding='utf-8')
            ws = wb.add_sheet('Battalion Three')

            row_num = 1
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = [
                '', '', '', '','','','','',
                title_doc,
                '','','','','','','','','','','','','','','','','',
                '']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            row_num = 3
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = ['FILE NAME', 'LAST NAME', 'FILE NUMBER', 'NIN','IPPS',
                'ACCOUNT NUMBER','TEL CONTACT','SEX','RANK','EDUCATION LEVEL','OTHER EDUCATION LEVEL','BANK','BRANCH','DEPARTEMENT','TITLE','STATUS','SHIFT','DATE OF ENLISTMENT','DATE OF TRANSFER','DATE OF PROMOTION','DATE OF BIRTH','ARMED','SECTION','LOCATION','ON LEAVE','LEAVE START DATE',
                'LEAVE END DATE', 'SPECIAL DUTY START DATE', 'SPECIAL DUTY END DATE']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            font_style = xlwt.XFStyle()
            
            rows = Battallion_three.objects.filter(department=query_parameter, on_leave=query_parameter2).values_list('first_name', 'last_name', 'file_number', 'nin', 
                'ipps','account_number','contact','sex','rank','education_level','other_education_level','bank','branch','department','title','status','shift','date_of_enlistment','date_of_transfer','date_of_promotion','date_of_birth','armed','section','location','on_leave','leave_start_date',
                'leave_end_date', 'special_duty_start_date', 'special_duty_end_date')

            for row in rows:
                row_num += 1

                for col_num in range(len(row)):
                    ws.write(row_num, col_num, str(row[col_num]), font_style)
            wb.save(response)

            return response
        else: 
            # print("You are not authorized to carry out this operation.")
            return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)
    except: 
        # print("You are not authorized to carry out this operation.")
        return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)

  
def export_excel_B_one_section_leave(request):
    try: 
        # to implement some security here 
        # print(request.method)
        print(request.GET['leave_type'] )
        print(request.GET['section'] )
        token = request.GET['unique']
        query_parameter = request.GET['section']
        query_parameter2 = request.GET['leave_type']
        title_doc = request.GET['title_doc']

        # title = 'BATTALION ONE ' + query_parameter + ' DATA'
        # print(title)

        if token == unique_token:
            response = HttpResponse(content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename=Expenses' + \
                 str(datetime.datetime.now())+'.xls'
            wb = xlwt.Workbook(encoding='utf-8')
            ws = wb.add_sheet('Battalion One')

            row_num = 1
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = [
                '', '', '', '','','','','',
                title_doc,
                '','','','','','','','','','','','','','','','','',
                '']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            row_num = 3
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = ['FILE NAME', 'LAST NAME', 'FILE NUMBER', 'NIN','IPPS',
                'ACCOUNT NUMBER','TEL CONTACT','SEX','RANK','EDUCATION LEVEL','OTHER EDUCATION LEVEL','BANK','BRANCH','DEPARTEMENT','TITLE','STATUS','SHIFT','DATE OF ENLISTMENT','DATE OF TRANSFER','DATE OF PROMOTION','DATE OF BIRTH','ARMED','SECTION','LOCATION','ON LEAVE','LEAVE START DATE',
                'LEAVE END DATE', 'SPECIAL DUTY START DATE', 'SPECIAL DUTY END DATE']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            font_style = xlwt.XFStyle()
            
            rows = Battallion_one.objects.filter(section=query_parameter, on_leave=query_parameter2).values_list('first_name', 'last_name', 'file_number', 'nin', 
                'ipps','account_number','contact','sex','rank','education_level','other_education_level','bank','branch','department','title','status','shift','date_of_enlistment','date_of_transfer','date_of_promotion','date_of_birth','armed','section','location','on_leave','leave_start_date',
                'leave_end_date', 'special_duty_start_date', 'special_duty_end_date')

            for row in rows:
                row_num += 1

                for col_num in range(len(row)):
                    ws.write(row_num, col_num, str(row[col_num]), font_style)
            wb.save(response)

            return response
        else: 
            # print("You are not authorized to carry out this operation.")
            return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)
    except: 
        # print("You are not authorized to carry out this operation.")
        return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)


def export_excel_B_five_department_status(request):
    try: 
        # to implement some security here 
        # print(request.method)
        # print(request.GET['status_type'] )
        # print(request.GET['department'] )
        token = request.GET['unique']
        query_parameter = request.GET['department']
        query_parameter2 = request.GET['status_type']
        title_doc = request.GET['title_doc']

        # title = 'BATTALION ONE ' + query_parameter + ' DATA'
        # print(title)

        if token == unique_token:
            response = HttpResponse(content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename=Expenses' + \
                 str(datetime.datetime.now())+'.xls'
            wb = xlwt.Workbook(encoding='utf-8')
            ws = wb.add_sheet('Battalion Five')

            row_num = 1
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = [
                '', '', '', '','','','','',
                title_doc,
                '','','','','','','','','','','','','','','','','',
                '']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            row_num = 3
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = ['FILE NAME', 'LAST NAME', 'FILE NUMBER', 'NIN','IPPS',
                'ACCOUNT NUMBER','TEL CONTACT','SEX','RANK','EDUCATION LEVEL','OTHER EDUCATION LEVEL','BANK','BRANCH','DEPARTEMENT','TITLE','STATUS','SHIFT','DATE OF ENLISTMENT','DATE OF TRANSFER','DATE OF PROMOTION','DATE OF BIRTH','ARMED','SECTION','LOCATION','ON LEAVE','LEAVE START DATE',
                'LEAVE END DATE', 'SPECIAL DUTY START DATE', 'SPECIAL DUTY END DATE']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            font_style = xlwt.XFStyle()
            
            rows = Battallion_five.objects.filter(department=query_parameter, status=query_parameter2).values_list('first_name', 'last_name', 'file_number', 'nin', 
                'ipps','account_number','contact','sex','rank','education_level','other_education_level','bank','branch','department','title','status','shift','date_of_enlistment','date_of_transfer','date_of_promotion','date_of_birth','armed','section','location','on_leave','leave_start_date',
                'leave_end_date', 'special_duty_start_date', 'special_duty_end_date')

            for row in rows:
                row_num += 1

                for col_num in range(len(row)):
                    ws.write(row_num, col_num, str(row[col_num]), font_style)
            wb.save(response)

            return response
        else: 
            # print("You are not authorized to carry out this operation.")
            return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)
    except: 
        # print("You are not authorized to carry out this operation.")
        return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)

def export_excel_B_four_department_status(request):
    try: 
        # to implement some security here 
        # print(request.method)
        # print(request.GET['status_type'] )
        # print(request.GET['department'] )
        token = request.GET['unique']
        query_parameter = request.GET['department']
        query_parameter2 = request.GET['status_type']
        title_doc = request.GET['title_doc']

        # title = 'BATTALION ONE ' + query_parameter + ' DATA'
        # print(title)

        if token == unique_token:
            response = HttpResponse(content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename=Expenses' + \
                 str(datetime.datetime.now())+'.xls'
            wb = xlwt.Workbook(encoding='utf-8')
            ws = wb.add_sheet('Battalion Four')

            row_num = 1
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = [
                '', '', '', '','','','','',
                title_doc,
                '','','','','','','','','','','','','','','','','',
                '']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            row_num = 3
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = ['FILE NAME', 'LAST NAME', 'FILE NUMBER', 'NIN','IPPS',
                'ACCOUNT NUMBER','TEL CONTACT','SEX','RANK','EDUCATION LEVEL','OTHER EDUCATION LEVEL','BANK','BRANCH','DEPARTEMENT','TITLE','STATUS','SHIFT','DATE OF ENLISTMENT','DATE OF TRANSFER','DATE OF PROMOTION','DATE OF BIRTH','ARMED','SECTION','LOCATION','ON LEAVE','LEAVE START DATE',
                'LEAVE END DATE', 'SPECIAL DUTY START DATE', 'SPECIAL DUTY END DATE']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            font_style = xlwt.XFStyle()
            
            rows = Battallion_four.objects.filter(department=query_parameter, status=query_parameter2).values_list('first_name', 'last_name', 'file_number', 'nin', 
                'ipps','account_number','contact','sex','rank','education_level','other_education_level','bank','branch','department','title','status','shift','date_of_enlistment','date_of_transfer','date_of_promotion','date_of_birth','armed','section','location','on_leave','leave_start_date',
                'leave_end_date', 'special_duty_start_date', 'special_duty_end_date')

            for row in rows:
                row_num += 1

                for col_num in range(len(row)):
                    ws.write(row_num, col_num, str(row[col_num]), font_style)
            wb.save(response)

            return response
        else: 
            # print("You are not authorized to carry out this operation.")
            return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)
    except: 
        # print("You are not authorized to carry out this operation.")
        return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)

def export_excel_B_three_department_status(request):
    try: 
        # to implement some security here 
        # print(request.method)
        # print(request.GET['status_type'] )
        # print(request.GET['department'] )
        token = request.GET['unique']
        query_parameter = request.GET['department']
        query_parameter2 = request.GET['status_type']
        title_doc = request.GET['title_doc']

        # title = 'BATTALION ONE ' + query_parameter + ' DATA'
        # print(title)

        if token == unique_token:
            response = HttpResponse(content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename=Expenses' + \
                 str(datetime.datetime.now())+'.xls'
            wb = xlwt.Workbook(encoding='utf-8')
            ws = wb.add_sheet('Battalion Three')

            row_num = 1
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = [
                '', '', '', '','','','','',
                title_doc,
                '','','','','','','','','','','','','','','','','',
                '']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            row_num = 3
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = ['FILE NAME', 'LAST NAME', 'FILE NUMBER', 'NIN','IPPS',
                'ACCOUNT NUMBER','TEL CONTACT','SEX','RANK','EDUCATION LEVEL','OTHER EDUCATION LEVEL','BANK','BRANCH','DEPARTEMENT','TITLE','STATUS','SHIFT','DATE OF ENLISTMENT','DATE OF TRANSFER','DATE OF PROMOTION','DATE OF BIRTH','ARMED','SECTION','LOCATION','ON LEAVE','LEAVE START DATE',
                'LEAVE END DATE', 'SPECIAL DUTY START DATE', 'SPECIAL DUTY END DATE']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            font_style = xlwt.XFStyle()
            
            rows = Battallion_three.objects.filter(department=query_parameter, status=query_parameter2).values_list('first_name', 'last_name', 'file_number', 'nin', 
                'ipps','account_number','contact','sex','rank','education_level','other_education_level','bank','branch','department','title','status','shift','date_of_enlistment','date_of_transfer','date_of_promotion','date_of_birth','armed','section','location','on_leave','leave_start_date',
                'leave_end_date', 'special_duty_start_date', 'special_duty_end_date')

            for row in rows:
                row_num += 1

                for col_num in range(len(row)):
                    ws.write(row_num, col_num, str(row[col_num]), font_style)
            wb.save(response)

            return response
        else: 
            # print("You are not authorized to carry out this operation.")
            return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)
    except: 
        # print("You are not authorized to carry out this operation.")
        return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)


def export_excel_B_one_section_status(request):
    try: 
        # to implement some security here 
        # print(request.method)
        # print(request.GET['status_type'] )
        # print(request.GET['section'] )
        token = request.GET['unique']
        query_parameter = request.GET['section']
        query_parameter2 = request.GET['status_type']
        title_doc = request.GET['title_doc']

        # title = 'BATTALION ONE ' + query_parameter + ' DATA'
        # print(title)

        if token == unique_token:
            response = HttpResponse(content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename=Expenses' + \
                 str(datetime.datetime.now())+'.xls'
            wb = xlwt.Workbook(encoding='utf-8')
            ws = wb.add_sheet('Battalion One')

            row_num = 1
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = [
                '', '', '', '','','','','',
                title_doc,
                '','','','','','','','','','','','','','','','','',
                '']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            row_num = 3
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = ['FILE NAME', 'LAST NAME', 'FILE NUMBER', 'NIN','IPPS',
                'ACCOUNT NUMBER','TEL CONTACT','SEX','RANK','EDUCATION LEVEL','OTHER EDUCATION LEVEL','BANK','BRANCH','DEPARTEMENT','TITLE','STATUS','SHIFT','DATE OF ENLISTMENT','DATE OF TRANSFER','DATE OF PROMOTION','DATE OF BIRTH','ARMED','SECTION','LOCATION','ON LEAVE','LEAVE START DATE',
                'LEAVE END DATE', 'SPECIAL DUTY START DATE', 'SPECIAL DUTY END DATE']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            font_style = xlwt.XFStyle()
            
            rows = Battallion_one.objects.filter(section=query_parameter, status=query_parameter2).values_list('first_name', 'last_name', 'file_number', 'nin', 
                'ipps','account_number','contact','sex','rank','education_level','other_education_level','bank','branch','department','title','status','shift','date_of_enlistment','date_of_transfer','date_of_promotion','date_of_birth','armed','section','location','on_leave','leave_start_date',
                'leave_end_date', 'special_duty_start_date', 'special_duty_end_date')

            for row in rows:
                row_num += 1

                for col_num in range(len(row)):
                    ws.write(row_num, col_num, str(row[col_num]), font_style)
            wb.save(response)

            return response
        else: 
            # print("You are not authorized to carry out this operation.")
            return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)
    except: 
        # print("You are not authorized to carry out this operation.")
        return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)

def export_excel_B_five_departmnt(request):
    try: 
        # to implement some security here 
        # print(request.method)
        # print(request.GET['department'])
        token = request.GET['unique']
        query_parameter = request.GET['department']
        title_doc = request.GET['title_doc']

        # title = 'BATTALION ONE ' + query_parameter + ' DATA'
        # print(title)

        if token == unique_token:
            response = HttpResponse(content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename=Expenses' + \
                 str(datetime.datetime.now())+'.xls'
            wb = xlwt.Workbook(encoding='utf-8')
            ws = wb.add_sheet('Battalion Five')

            row_num = 1
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = [
                '', '', '', '','','','','',
                title_doc,
                '','','','','','','','','','','','','','','','','',
                '']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            row_num = 3
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = ['FILE NAME', 'LAST NAME', 'FILE NUMBER', 'NIN','IPPS',
                'ACCOUNT NUMBER','TEL CONTACT','SEX','RANK','EDUCATION LEVEL','OTHER EDUCATION LEVEL','BANK','BRANCH','DEPARTEMENT','TITLE','STATUS','SHIFT','DATE OF ENLISTMENT','DATE OF TRANSFER','DATE OF PROMOTION','DATE OF BIRTH','ARMED','SECTION','LOCATION','ON LEAVE','LEAVE START DATE',
                'LEAVE END DATE', 'SPECIAL DUTY START DATE', 'SPECIAL DUTY END DATE']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            font_style = xlwt.XFStyle()
            
            rows = Battallion_five.objects.filter(department=query_parameter).values_list('first_name', 'last_name', 'file_number', 'nin', 
                'ipps','account_number','contact','sex','rank','education_level','other_education_level','bank','branch','department','title','status','shift','date_of_enlistment','date_of_transfer','date_of_promotion','date_of_birth','armed','section','location','on_leave','leave_start_date',
                'leave_end_date', 'special_duty_start_date', 'special_duty_end_date')

            for row in rows:
                row_num += 1

                for col_num in range(len(row)):
                    ws.write(row_num, col_num, str(row[col_num]), font_style)
            wb.save(response)

            return response
        else: 
            # print("You are not authorized to carry out this operation.")
            return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)
    except: 
        # print("You are not authorized to carry out this operation.")
        return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)

def export_excel_B_four_departmnt(request):
    try: 
        # to implement some security here 
        # print(request.method)
        # print(request.GET['department'])
        token = request.GET['unique']
        query_parameter = request.GET['department']
        title_doc = request.GET['title_doc']

        # title = 'BATTALION ONE ' + query_parameter + ' DATA'
        # print(title)

        if token == unique_token:
            response = HttpResponse(content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename=Expenses' + \
                 str(datetime.datetime.now())+'.xls'
            wb = xlwt.Workbook(encoding='utf-8')
            ws = wb.add_sheet('Battalion Four')

            row_num = 1
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = [
                '', '', '', '','','','','',
                title_doc,
                '','','','','','','','','','','','','','','','','',
                '']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            row_num = 3
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = ['FILE NAME', 'LAST NAME', 'FILE NUMBER', 'NIN','IPPS',
                'ACCOUNT NUMBER','TEL CONTACT','SEX','RANK','EDUCATION LEVEL','OTHER EDUCATION LEVEL','BANK','BRANCH','DEPARTEMENT','TITLE','STATUS','SHIFT','DATE OF ENLISTMENT','DATE OF TRANSFER','DATE OF PROMOTION','DATE OF BIRTH','ARMED','SECTION','LOCATION','ON LEAVE','LEAVE START DATE',
                'LEAVE END DATE', 'SPECIAL DUTY START DATE', 'SPECIAL DUTY END DATE']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            font_style = xlwt.XFStyle()
            
            rows = Battallion_four.objects.filter(department=query_parameter).values_list('first_name', 'last_name', 'file_number', 'nin', 
                'ipps','account_number','contact','sex','rank','education_level','other_education_level','bank','branch','department','title','status','shift','date_of_enlistment','date_of_transfer','date_of_promotion','date_of_birth','armed','section','location','on_leave','leave_start_date',
                'leave_end_date', 'special_duty_start_date', 'special_duty_end_date')

            for row in rows:
                row_num += 1

                for col_num in range(len(row)):
                    ws.write(row_num, col_num, str(row[col_num]), font_style)
            wb.save(response)

            return response
        else: 
            # print("You are not authorized to carry out this operation.")
            return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)
    except: 
        # print("You are not authorized to carry out this operation.")
        return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)

def export_excel_B_three_departmnt(request):
    try: 
        # to implement some security here 
        # print(request.method)
        # print(request.GET['department'])
        token = request.GET['unique']
        query_parameter = request.GET['department']
        title_doc = request.GET['title_doc']

        # title = 'BATTALION ONE ' + query_parameter + ' DATA'
        # print(title)

        if token == unique_token:
            response = HttpResponse(content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename=Expenses' + \
                 str(datetime.datetime.now())+'.xls'
            wb = xlwt.Workbook(encoding='utf-8')
            ws = wb.add_sheet('Battalion Three')

            row_num = 1
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = [
                '', '', '', '','','','','',
                title_doc,
                '','','','','','','','','','','','','','','','','',
                '']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            row_num = 3
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = ['FILE NAME', 'LAST NAME', 'FILE NUMBER', 'NIN','IPPS',
                'ACCOUNT NUMBER','TEL CONTACT','SEX','RANK','EDUCATION LEVEL','OTHER EDUCATION LEVEL','BANK','BRANCH','DEPARTEMENT','TITLE','STATUS','SHIFT','DATE OF ENLISTMENT','DATE OF TRANSFER','DATE OF PROMOTION','DATE OF BIRTH','ARMED','SECTION','LOCATION','ON LEAVE','LEAVE START DATE',
                'LEAVE END DATE', 'SPECIAL DUTY START DATE', 'SPECIAL DUTY END DATE']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            font_style = xlwt.XFStyle()
            
            rows = Battallion_three.objects.filter(department=query_parameter).values_list('first_name', 'last_name', 'file_number', 'nin', 
                'ipps','account_number','contact','sex','rank','education_level','other_education_level','bank','branch','department','title','status','shift','date_of_enlistment','date_of_transfer','date_of_promotion','date_of_birth','armed','section','location','on_leave','leave_start_date',
                'leave_end_date', 'special_duty_start_date', 'special_duty_end_date')

            for row in rows:
                row_num += 1

                for col_num in range(len(row)):
                    ws.write(row_num, col_num, str(row[col_num]), font_style)
            wb.save(response)

            return response
        else: 
            # print("You are not authorized to carry out this operation.")
            return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)
    except: 
        # print("You are not authorized to carry out this operation.")
        return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)


def export_excel_B_one_section(request):
    try: 
        # to implement some security here 
        # print(request.method)
        # print(request.GET['section'] )
        token = request.GET['unique']
        query_parameter = request.GET['section']
        title_doc = request.GET['title_doc']

        # title = 'BATTALION ONE ' + query_parameter + ' DATA'
        # print(title)

        if token == unique_token:
            response = HttpResponse(content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename=Expenses' + \
                 str(datetime.datetime.now())+'.xls'
            wb = xlwt.Workbook(encoding='utf-8')
            ws = wb.add_sheet('Battalion One')

            row_num = 1
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = [
                '', '', '', '','','','','',
                title_doc,
                '','','','','','','','','','','','','','','','','',
                '']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            row_num = 3
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = ['FILE NAME', 'LAST NAME', 'FILE NUMBER', 'NIN','IPPS',
                'ACCOUNT NUMBER','TEL CONTACT','SEX','RANK','EDUCATION LEVEL','OTHER EDUCATION LEVEL','BANK','BRANCH','DEPARTEMENT','TITLE','STATUS','SHIFT','DATE OF ENLISTMENT','DATE OF TRANSFER','DATE OF PROMOTION','DATE OF BIRTH','ARMED','SECTION','LOCATION','ON LEAVE','LEAVE START DATE',
                'LEAVE END DATE', 'SPECIAL DUTY START DATE', 'SPECIAL DUTY END DATE']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            font_style = xlwt.XFStyle()
            
            rows = Battallion_one.objects.filter(section=query_parameter).values_list('first_name', 'last_name', 'file_number', 'nin', 
                'ipps','account_number','contact','sex','rank','education_level','other_education_level','bank','branch','department','title','status','shift','date_of_enlistment','date_of_transfer','date_of_promotion','date_of_birth','armed','section','location','on_leave','leave_start_date',
                'leave_end_date', 'special_duty_start_date', 'special_duty_end_date')

            for row in rows:
                row_num += 1

                for col_num in range(len(row)):
                    ws.write(row_num, col_num, str(row[col_num]), font_style)
            wb.save(response)

            return response
        else: 
            # print("You are not authorized to carry out this operation.")
            return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)
    except: 
        # print("You are not authorized to carry out this operation.")
        return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)


def export_excel_B_five_leave(request):
    try: 

        token = request.GET['unique']
        query_parameter = request.GET['leave_type']
        title_doc = request.GET['title_doc']


        if token == unique_token:
            response = HttpResponse(content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename=Expenses' + \
                 str(datetime.datetime.now())+'.xls'
            wb = xlwt.Workbook(encoding='utf-8')
            ws = wb.add_sheet('Battalion Five')

            row_num = 1
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = [
                '', '', '', '','','','','',
                title_doc,
                '','','','','','','','','','','','','','','','','',
                '']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            row_num = 3
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = ['FILE NAME', 'LAST NAME', 'FILE NUMBER', 'NIN','IPPS',
                'ACCOUNT NUMBER','TEL CONTACT','SEX','RANK','EDUCATION LEVEL','OTHER EDUCATION LEVEL','BANK','BRANCH','DEPARTEMENT','TITLE','STATUS','SHIFT','DATE OF ENLISTMENT','DATE OF TRANSFER','DATE OF PROMOTION','DATE OF BIRTH','ARMED','SECTION','LOCATION','ON LEAVE','LEAVE START DATE',
                'LEAVE END DATE', 'SPECIAL DUTY START DATE', 'SPECIAL DUTY END DATE']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            font_style = xlwt.XFStyle()
            
            rows = Battallion_five.objects.filter(on_leave=query_parameter).values_list('first_name', 'last_name', 'file_number', 'nin', 
                'ipps','account_number','contact','sex','rank','education_level','other_education_level','bank','branch','department','title','status','shift','date_of_enlistment','date_of_transfer','date_of_promotion','date_of_birth','armed','section','location','on_leave','leave_start_date',
                'leave_end_date', 'special_duty_start_date', 'special_duty_end_date')

            for row in rows:
                row_num += 1

                for col_num in range(len(row)):
                    ws.write(row_num, col_num, str(row[col_num]), font_style)
            wb.save(response)

            return response
        else: 
            # print("You are not authorized to carry out this operation.")
            return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)
    except: 
        # print("You are not authorized to carry out this operation.")
        return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)

def export_excel_B_four_leave(request):
    try: 

        token = request.GET['unique']
        query_parameter = request.GET['leave_type']
        title_doc = request.GET['title_doc']


        if token == unique_token:
            response = HttpResponse(content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename=Expenses' + \
                 str(datetime.datetime.now())+'.xls'
            wb = xlwt.Workbook(encoding='utf-8')
            ws = wb.add_sheet('Battalion Four')

            row_num = 1
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = [
                '', '', '', '','','','','',
                title_doc,
                '','','','','','','','','','','','','','','','','',
                '']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            row_num = 3
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = ['FILE NAME', 'LAST NAME', 'FILE NUMBER', 'NIN','IPPS',
                'ACCOUNT NUMBER','TEL CONTACT','SEX','RANK','EDUCATION LEVEL','OTHER EDUCATION LEVEL','BANK','BRANCH','DEPARTEMENT','TITLE','STATUS','SHIFT','DATE OF ENLISTMENT','DATE OF TRANSFER','DATE OF PROMOTION','DATE OF BIRTH','ARMED','SECTION','LOCATION','ON LEAVE','LEAVE START DATE',
                'LEAVE END DATE', 'SPECIAL DUTY START DATE', 'SPECIAL DUTY END DATE']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            font_style = xlwt.XFStyle()
            
            rows = Battallion_four.objects.filter(on_leave=query_parameter).values_list('first_name', 'last_name', 'file_number', 'nin', 
                'ipps','account_number','contact','sex','rank','education_level','other_education_level','bank','branch','department','title','status','shift','date_of_enlistment','date_of_transfer','date_of_promotion','date_of_birth','armed','section','location','on_leave','leave_start_date',
                'leave_end_date', 'special_duty_start_date', 'special_duty_end_date')

            for row in rows:
                row_num += 1

                for col_num in range(len(row)):
                    ws.write(row_num, col_num, str(row[col_num]), font_style)
            wb.save(response)

            return response
        else: 
            # print("You are not authorized to carry out this operation.")
            return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)
    except: 
        # print("You are not authorized to carry out this operation.")
        return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)


def export_excel_B_three_leave(request):
    try: 

        token = request.GET['unique']
        query_parameter = request.GET['leave_type']
        title_doc = request.GET['title_doc']


        if token == unique_token:
            response = HttpResponse(content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename=Expenses' + \
                 str(datetime.datetime.now())+'.xls'
            wb = xlwt.Workbook(encoding='utf-8')
            ws = wb.add_sheet('Battalion Three')

            row_num = 1
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = [
                '', '', '', '','','','','',
                title_doc,
                '','','','','','','','','','','','','','','','','',
                '']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            row_num = 3
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = ['FILE NAME', 'LAST NAME', 'FILE NUMBER', 'NIN','IPPS',
                'ACCOUNT NUMBER','TEL CONTACT','SEX','RANK','EDUCATION LEVEL','OTHER EDUCATION LEVEL','BANK','BRANCH','DEPARTEMENT','TITLE','STATUS','SHIFT','DATE OF ENLISTMENT','DATE OF TRANSFER','DATE OF PROMOTION','DATE OF BIRTH','ARMED','SECTION','LOCATION','ON LEAVE','LEAVE START DATE',
                'LEAVE END DATE', 'SPECIAL DUTY START DATE', 'SPECIAL DUTY END DATE']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            font_style = xlwt.XFStyle()
            
            rows = Battallion_three.objects.filter(on_leave=query_parameter).values_list('first_name', 'last_name', 'file_number', 'nin', 
                'ipps','account_number','contact','sex','rank','education_level','other_education_level','bank','branch','department','title','status','shift','date_of_enlistment','date_of_transfer','date_of_promotion','date_of_birth','armed','section','location','on_leave','leave_start_date',
                'leave_end_date', 'special_duty_start_date', 'special_duty_end_date')

            for row in rows:
                row_num += 1

                for col_num in range(len(row)):
                    ws.write(row_num, col_num, str(row[col_num]), font_style)
            wb.save(response)

            return response
        else: 
            # print("You are not authorized to carry out this operation.")
            return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)
    except: 
        # print("You are not authorized to carry out this operation.")
        return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)


def export_excel_B_one_leave(request):
    try: 
        # to implement some security here 
        # print(request.method)
        # print(request.GET['leave_type'] )
        token = request.GET['unique']
        query_parameter = request.GET['leave_type']
        title_doc = request.GET['title_doc']

        # title = 'BATTALION ONE ' + query_parameter + ' DATA'
        # print(title)

        if token == unique_token:
            response = HttpResponse(content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename=Expenses' + \
                 str(datetime.datetime.now())+'.xls'
            wb = xlwt.Workbook(encoding='utf-8')
            ws = wb.add_sheet('Battalion One')

            row_num = 1
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = [
                '', '', '', '','','','','',
                title_doc,
                '','','','','','','','','','','','','','','','','',
                '']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            row_num = 3
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = ['FILE NAME', 'LAST NAME', 'FILE NUMBER', 'NIN','IPPS',
                'ACCOUNT NUMBER','TEL CONTACT','SEX','RANK','EDUCATION LEVEL','OTHER EDUCATION LEVEL','BANK','BRANCH','DEPARTEMENT','TITLE','STATUS','SHIFT','DATE OF ENLISTMENT','DATE OF TRANSFER','DATE OF PROMOTION','DATE OF BIRTH','ARMED','SECTION','LOCATION','ON LEAVE','LEAVE START DATE',
                'LEAVE END DATE', 'SPECIAL DUTY START DATE', 'SPECIAL DUTY END DATE']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            font_style = xlwt.XFStyle()
            
            rows = Battallion_one.objects.filter(on_leave=query_parameter).values_list('first_name', 'last_name', 'file_number', 'nin', 
                'ipps','account_number','contact','sex','rank','education_level','other_education_level','bank','branch','department','title','status','shift','date_of_enlistment','date_of_transfer','date_of_promotion','date_of_birth','armed','section','location','on_leave','leave_start_date',
                'leave_end_date', 'special_duty_start_date', 'special_duty_end_date')

            for row in rows:
                row_num += 1

                for col_num in range(len(row)):
                    ws.write(row_num, col_num, str(row[col_num]), font_style)
            wb.save(response)

            return response
        else: 
            # print("You are not authorized to carry out this operation.")
            return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)
    except: 
        # print("You are not authorized to carry out this operation.")
        return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)


def export_excel_B_two_leave(request):
    try: 
        # to implement some security here 
        # print(request.method)
        print(request.GET['leave_type'] )
        token = request.GET['unique']
        query_parameter = request.GET['leave_type']
        title_doc = request.GET['title_doc']

        # title = 'BATTALION TWO ' + query_parameter + ' DATA'
        # print(title)

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
                title_doc,
                '','','','','','','','','','','','','','','','','',
                '']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            row_num = 3
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = ['FILE NAME', 'LAST NAME', 'FILE NUMBER', 'NIN','IPPS',
                'ACCOUNT NUMBER','TEL CONTACT','SEX','RANK','EDUCATION LEVEL','OTHER EDUCATION LEVEL','BANK','BRANCH','DEPARTEMENT','TITLE','STATUS','SHIFT','DATE OF ENLISTMENT','DATE OF TRANSFER','DATE OF PROMOTION','DATE OF BIRTH','ARMED','SECTION','LOCATION','ON LEAVE','LEAVE START DATE',
                'LEAVE END DATE', 'SPECIAL DUTY START DATE', 'SPECIAL DUTY END DATE']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            font_style = xlwt.XFStyle()
            
            rows = Battallion_two.objects.filter(on_leave=query_parameter).values_list('first_name', 'last_name', 'file_number', 'nin', 
                'ipps','account_number','contact','sex','rank','education_level','other_education_level','bank','branch','department','title','status','shift','date_of_enlistment','date_of_transfer','date_of_promotion','date_of_birth','armed','section','location','on_leave','leave_start_date',
                'leave_end_date', 'special_duty_start_date', 'special_duty_end_date')

            for row in rows:
                row_num += 1

                for col_num in range(len(row)):
                    ws.write(row_num, col_num, str(row[col_num]), font_style)
            wb.save(response)

            return response
        else: 
            # print("You are not authorized to carry out this operation.")
            return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)
    except: 
        # print("You are not authorized to carry out this operation.")
        return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)


def export_excel_B_five_status(request):
    try: 
        token = request.GET['unique']
        query_parameter = request.GET['status_type']
        title_doc = request.GET['title_doc']

        if token == unique_token:
            response = HttpResponse(content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename=Expenses' + \
                 str(datetime.datetime.now())+'.xls'
            wb = xlwt.Workbook(encoding='utf-8')
            ws = wb.add_sheet('Battalion Five')

            row_num = 1
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = [
                '', '', '', '','','','','',
                title_doc,
                '','','','','','','','','','','','','','','','','',
                '']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            row_num = 3
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = ['FILE NAME', 'LAST NAME', 'FILE NUMBER', 'NIN','IPPS',
                'ACCOUNT NUMBER','TEL CONTACT','SEX','RANK','EDUCATION LEVEL','OTHER EDUCATION LEVEL','BANK','BRANCH','DEPARTEMENT','TITLE','STATUS','SHIFT','DATE OF ENLISTMENT','DATE OF TRANSFER','DATE OF PROMOTION','DATE OF BIRTH','ARMED','SECTION','LOCATION','ON LEAVE','LEAVE START DATE',
                'LEAVE END DATE', 'SPECIAL DUTY START DATE', 'SPECIAL DUTY END DATE']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            font_style = xlwt.XFStyle()
            
            rows = Battallion_five.objects.filter(status=query_parameter).values_list('first_name', 'last_name', 'file_number', 'nin', 
                'ipps','account_number','contact','sex','rank','education_level','other_education_level','bank','branch','department','title','status','shift','date_of_enlistment','date_of_transfer','date_of_promotion','date_of_birth','armed','section','location','on_leave','leave_start_date',
                'leave_end_date', 'special_duty_start_date', 'special_duty_end_date')

            for row in rows:
                row_num += 1

                for col_num in range(len(row)):
                    ws.write(row_num, col_num, str(row[col_num]), font_style)
            wb.save(response)

            return response
        else: 
            # print("You are not authorized to carry out this operation.")
            return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)
    except: 
        # print("You are not authorized to carry out this operation.")
        return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)

def export_excel_B_four_status(request):
    try: 
        token = request.GET['unique']
        query_parameter = request.GET['status_type']
        title_doc = request.GET['title_doc']

        if token == unique_token:
            response = HttpResponse(content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename=Expenses' + \
                 str(datetime.datetime.now())+'.xls'
            wb = xlwt.Workbook(encoding='utf-8')
            ws = wb.add_sheet('Battalion Four')

            row_num = 1
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = [
                '', '', '', '','','','','',
                title_doc,
                '','','','','','','','','','','','','','','','','',
                '']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            row_num = 3
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = ['FILE NAME', 'LAST NAME', 'FILE NUMBER', 'NIN','IPPS',
                'ACCOUNT NUMBER','TEL CONTACT','SEX','RANK','EDUCATION LEVEL','OTHER EDUCATION LEVEL','BANK','BRANCH','DEPARTEMENT','TITLE','STATUS','SHIFT','DATE OF ENLISTMENT','DATE OF TRANSFER','DATE OF PROMOTION','DATE OF BIRTH','ARMED','SECTION','LOCATION','ON LEAVE','LEAVE START DATE',
                'LEAVE END DATE', 'SPECIAL DUTY START DATE', 'SPECIAL DUTY END DATE']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            font_style = xlwt.XFStyle()
            
            rows = Battallion_four.objects.filter(status=query_parameter).values_list('first_name', 'last_name', 'file_number', 'nin', 
                'ipps','account_number','contact','sex','rank','education_level','other_education_level','bank','branch','department','title','status','shift','date_of_enlistment','date_of_transfer','date_of_promotion','date_of_birth','armed','section','location','on_leave','leave_start_date',
                'leave_end_date', 'special_duty_start_date', 'special_duty_end_date')

            for row in rows:
                row_num += 1

                for col_num in range(len(row)):
                    ws.write(row_num, col_num, str(row[col_num]), font_style)
            wb.save(response)

            return response
        else: 
            # print("You are not authorized to carry out this operation.")
            return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)
    except: 
        # print("You are not authorized to carry out this operation.")
        return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)

def export_excel_B_three_status(request):
    try: 
        token = request.GET['unique']
        query_parameter = request.GET['status_type']
        title_doc = request.GET['title_doc']

        if token == unique_token:
            response = HttpResponse(content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename=Expenses' + \
                 str(datetime.datetime.now())+'.xls'
            wb = xlwt.Workbook(encoding='utf-8')
            ws = wb.add_sheet('Battalion Three')

            row_num = 1
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = [
                '', '', '', '','','','','',
                title_doc,
                '','','','','','','','','','','','','','','','','',
                '']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            row_num = 3
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = ['FILE NAME', 'LAST NAME', 'FILE NUMBER', 'NIN','IPPS',
                'ACCOUNT NUMBER','TEL CONTACT','SEX','RANK','EDUCATION LEVEL','OTHER EDUCATION LEVEL','BANK','BRANCH','DEPARTEMENT','TITLE','STATUS','SHIFT','DATE OF ENLISTMENT','DATE OF TRANSFER','DATE OF PROMOTION','DATE OF BIRTH','ARMED','SECTION','LOCATION','ON LEAVE','LEAVE START DATE',
                'LEAVE END DATE', 'SPECIAL DUTY START DATE', 'SPECIAL DUTY END DATE']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            font_style = xlwt.XFStyle()
            
            rows = Battallion_three.objects.filter(status=query_parameter).values_list('first_name', 'last_name', 'file_number', 'nin', 
                'ipps','account_number','contact','sex','rank','education_level','other_education_level','bank','branch','department','title','status','shift','date_of_enlistment','date_of_transfer','date_of_promotion','date_of_birth','armed','section','location','on_leave','leave_start_date',
                'leave_end_date', 'special_duty_start_date', 'special_duty_end_date')

            for row in rows:
                row_num += 1

                for col_num in range(len(row)):
                    ws.write(row_num, col_num, str(row[col_num]), font_style)
            wb.save(response)

            return response
        else: 
            # print("You are not authorized to carry out this operation.")
            return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)
    except: 
        # print("You are not authorized to carry out this operation.")
        return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)


def export_excel_B_one_status(request):
    try: 
        # to implement some security here 
        # print(request.method)
        # print(request.GET['status_type'] )
        token = request.GET['unique']
        query_parameter = request.GET['status_type']
        title_doc = request.GET['title_doc']

        # title = 'BATTALION ONE ' + query_parameter + ' DATA'
        # print(title)

        if token == unique_token:
            response = HttpResponse(content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename=Expenses' + \
                 str(datetime.datetime.now())+'.xls'
            wb = xlwt.Workbook(encoding='utf-8')
            ws = wb.add_sheet('Battalion One')

            row_num = 1
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = [
                '', '', '', '','','','','',
                title_doc,
                '','','','','','','','','','','','','','','','','',
                '']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            row_num = 3
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = ['FILE NAME', 'LAST NAME', 'FILE NUMBER', 'NIN','IPPS',
                'ACCOUNT NUMBER','TEL CONTACT','SEX','RANK','EDUCATION LEVEL','OTHER EDUCATION LEVEL','BANK','BRANCH','DEPARTEMENT','TITLE','STATUS','SHIFT','DATE OF ENLISTMENT','DATE OF TRANSFER','DATE OF PROMOTION','DATE OF BIRTH','ARMED','SECTION','LOCATION','ON LEAVE','LEAVE START DATE',
                'LEAVE END DATE', 'SPECIAL DUTY START DATE', 'SPECIAL DUTY END DATE']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            font_style = xlwt.XFStyle()
            
            rows = Battallion_one.objects.filter(status=query_parameter).values_list('first_name', 'last_name', 'file_number', 'nin', 
                'ipps','account_number','contact','sex','rank','education_level','other_education_level','bank','branch','department','title','status','shift','date_of_enlistment','date_of_transfer','date_of_promotion','date_of_birth','armed','section','location','on_leave','leave_start_date',
                'leave_end_date', 'special_duty_start_date', 'special_duty_end_date')

            for row in rows:
                row_num += 1

                for col_num in range(len(row)):
                    ws.write(row_num, col_num, str(row[col_num]), font_style)
            wb.save(response)

            return response
        else: 
            # print("You are not authorized to carry out this operation.")
            return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)
    except: 
        # print("You are not authorized to carry out this operation.")
        return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)


def export_excel_B_two_status(request):
    try: 
        # to implement some security here 
        # print(request.method)
        print(request.GET['status_type'] )
        token = request.GET['unique']
        query_parameter = request.GET['status_type']
        title_doc = request.GET['title_doc']

        # title = 'BATTALION TWO ' + query_parameter + ' DATA'
        # print(title)

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
                title_doc,
                '','','','','','','','','','','','','','','','','',
                '']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            row_num = 3
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = ['FILE NAME', 'LAST NAME', 'FILE NUMBER', 'NIN','IPPS',
                'ACCOUNT NUMBER','TEL CONTACT','SEX','RANK','EDUCATION LEVEL','OTHER EDUCATION LEVEL','BANK','BRANCH','DEPARTEMENT','TITLE','STATUS','SHIFT',
                'DATE OF ENLISTMENT','DATE OF TRANSFER','DATE OF PROMOTION','DATE OF BIRTH','ARMED','SECTION','LOCATION','ON LEAVE','LEAVE START DATE',
                'LEAVE END DATE', 'SPECIAL DUTY START DATE', 'SPECIAL DUTY END DATE']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            font_style = xlwt.XFStyle()
            
            rows = Battallion_two.objects.filter(status=query_parameter).values_list('first_name', 'last_name', 'file_number', 'nin', 
                'ipps','account_number','contact','sex','rank','education_level','other_education_level','bank','branch','department','title','status','shift',
                'date_of_enlistment','date_of_transfer','date_of_promotion','date_of_birth','armed','section','location','on_leave','leave_start_date',
                'leave_end_date', 'special_duty_start_date', 'special_duty_end_date')

            for row in rows:
                row_num += 1

                for col_num in range(len(row)):
                    ws.write(row_num, col_num, str(row[col_num]), font_style)
            wb.save(response)

            return response
        else: 
            # print("You are not authorized to carry out this operation.")
            return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)
    except: 
        # print("You are not authorized to carry out this operation.")
        return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)


def export_excel_B_five(request):
    try: 
        # to implement some security here 
        # print(request.method)
        # print(request.GET['unique'] )
        token = request.GET['unique']
        title_doc = request.GET['title_doc']

        if token == unique_token:
            response = HttpResponse(content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename=Expenses' + \
                 str(datetime.datetime.now())+'.xls'
            wb = xlwt.Workbook(encoding='utf-8')
            ws = wb.add_sheet('Battalion Five')

            row_num = 1
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = [
                '', '', '', '','','','','',
                title_doc,
                '','','','','','','','','','','','','','','','','',
                '']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            row_num = 3
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = ['FILE NAME', 'LAST NAME', 'FILE NUMBER', 'NIN','IPPS',
                'ACCOUNT NUMBER','TEL CONTACT','SEX','RANK','EDUCATION LEVEL','OTHER EDUCATION LEVEL','BANK','BRANCH','DEPARTEMENT',
                'TITLE','STATUS','SHIFT','DATE OF ENLISTMENT','DATE OF TRANSFER','DATE OF PROMOTION','DATE OF BIRTH','ARMED','SECTION','LOCATION','ON LEAVE','LEAVE START DATE',
                'LEAVE END DATE', 'SPECIAL DUTY START DATE', 'SPECIAL DUTY END DATE']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            font_style = xlwt.XFStyle()
            
            # for reports with specific Sections
            # query_parameter = "UN Women"
            # rows = Battallion_one.objects.filter(section=query_parameter).values_list('first_name', 'last_name', 'file_number', 'nin', 

            rows = Battallion_five.objects.filter().values_list('first_name', 'last_name', 'file_number', 'nin', 
                'ipps','account_number','contact','sex','rank','education_level','other_education_level','bank','branch','department',
                'title','status','shift','date_of_enlistment','date_of_transfer','date_of_promotion','date_of_birth','armed','section','location','on_leave','leave_start_date',
                'leave_end_date', 'special_duty_start_date', 'special_duty_end_date')

            for row in rows:
                row_num += 1

                for col_num in range(len(row)):
                    ws.write(row_num, col_num, str(row[col_num]), font_style)
            wb.save(response)

            return response
        else: 
            # print("You are not authorized to carry out this operation.")
            return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)
    except: 
        # print("You are not authorized to carry out this operation.")
        return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)

def export_excel_B_four(request):
    try: 
        # to implement some security here 
        # print(request.method)
        # print(request.GET['unique'] )
        token = request.GET['unique']
        title_doc = request.GET['title_doc']

        if token == unique_token:
            response = HttpResponse(content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename=Expenses' + \
                 str(datetime.datetime.now())+'.xls'
            wb = xlwt.Workbook(encoding='utf-8')
            ws = wb.add_sheet('Battalion Four')

            row_num = 1
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = [
                '', '', '', '','','','','',
                title_doc,
                '','','','','','','','','','','','','','','','','',
                '']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            row_num = 3
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = ['FILE NAME', 'LAST NAME', 'FILE NUMBER', 'NIN','IPPS',
                'ACCOUNT NUMBER','TEL CONTACT','SEX','RANK','EDUCATION LEVEL','OTHER EDUCATION LEVEL','BANK','BRANCH','DEPARTEMENT',
                'TITLE','STATUS','SHIFT','DATE OF ENLISTMENT','DATE OF TRANSFER','DATE OF PROMOTION','DATE OF BIRTH','ARMED','SECTION','LOCATION','ON LEAVE','LEAVE START DATE',
                'LEAVE END DATE', 'SPECIAL DUTY START DATE', 'SPECIAL DUTY END DATE']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            font_style = xlwt.XFStyle()
            
            # for reports with specific Sections
            # query_parameter = "UN Women"
            # rows = Battallion_one.objects.filter(section=query_parameter).values_list('first_name', 'last_name', 'file_number', 'nin', 

            rows = Battallion_four.objects.filter().values_list('first_name', 'last_name', 'file_number', 'nin', 
                'ipps','account_number','contact','sex','rank','education_level','other_education_level','bank','branch','department',
                'title','status','shift','date_of_enlistment','date_of_transfer','date_of_promotion','date_of_birth','armed','section','location','on_leave','leave_start_date',
                'leave_end_date', 'special_duty_start_date', 'special_duty_end_date')

            for row in rows:
                row_num += 1

                for col_num in range(len(row)):
                    ws.write(row_num, col_num, str(row[col_num]), font_style)
            wb.save(response)

            return response
        else: 
            # print("You are not authorized to carry out this operation.")
            return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)
    except: 
        # print("You are not authorized to carry out this operation.")
        return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)

def export_excel_B_three(request):
    try: 
        # to implement some security here 
        # print(request.method)
        print(request.GET['unique'] )
        token = request.GET['unique']
        title_doc = request.GET['title_doc']

        if token == unique_token:
            response = HttpResponse(content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename=Expenses' + \
                 str(datetime.datetime.now())+'.xls'
            wb = xlwt.Workbook(encoding='utf-8')
            ws = wb.add_sheet('Battalion Three')

            row_num = 1
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = [
                '', '', '', '','','','','',
                title_doc,
                '','','','','','','','','','','','','','','','','',
                '']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            row_num = 3
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = ['FILE NAME', 'LAST NAME', 'FILE NUMBER', 'NIN','IPPS',
                'ACCOUNT NUMBER','TEL CONTACT','SEX','RANK','EDUCATION LEVEL','OTHER EDUCATION LEVEL','BANK','BRANCH','DEPARTEMENT',
                'TITLE','STATUS','SHIFT','DATE OF ENLISTMENT','DATE OF TRANSFER','DATE OF PROMOTION','DATE OF BIRTH','ARMED','SECTION','LOCATION','ON LEAVE','LEAVE START DATE',
                'LEAVE END DATE', 'SPECIAL DUTY START DATE', 'SPECIAL DUTY END DATE']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            font_style = xlwt.XFStyle()
            
            # for reports with specific Sections
            # query_parameter = "UN Women"
            # rows = Battallion_one.objects.filter(section=query_parameter).values_list('first_name', 'last_name', 'file_number', 'nin', 

            rows = Battallion_three.objects.filter().values_list('first_name', 'last_name', 'file_number', 'nin', 
                'ipps','account_number','contact','sex','rank','education_level','other_education_level','bank','branch','department',
                'title','status','shift','date_of_enlistment','date_of_transfer','date_of_promotion','date_of_birth','armed','section','location','on_leave','leave_start_date',
                'leave_end_date', 'special_duty_start_date', 'special_duty_end_date')

            for row in rows:
                row_num += 1

                for col_num in range(len(row)):
                    ws.write(row_num, col_num, str(row[col_num]), font_style)
            wb.save(response)

            return response
        else: 
            # print("You are not authorized to carry out this operation.")
            return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)
    except: 
        # print("You are not authorized to carry out this operation.")
        return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)



def export_excel_B_one(request):
    try: 
        # to implement some security here 
        # print(request.method)
        print(request.GET['unique'] )
        token = request.GET['unique']
        title_doc = request.GET['title_doc']

        if token == unique_token:
            response = HttpResponse(content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename=Expenses' + \
                 str(datetime.datetime.now())+'.xls'
            wb = xlwt.Workbook(encoding='utf-8')
            ws = wb.add_sheet('Battalion One')

            row_num = 1
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = [
                '', '', '', '','','','','',
                title_doc,
                '','','','','','','','','','','','','','','','','',
                '']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            row_num = 3
            font_style = xlwt.XFStyle()
            font_style.font.bold = True

            columns = ['FILE NAME', 'LAST NAME', 'FILE NUMBER', 'NIN','IPPS',
                'ACCOUNT NUMBER','TEL CONTACT','SEX','RANK','EDUCATION LEVEL','OTHER EDUCATION LEVEL','BANK','BRANCH','DEPARTEMENT',
                'TITLE','STATUS','SHIFT','DATE OF ENLISTMENT','DATE OF TRANSFER','DATE OF PROMOTION','DATE OF BIRTH','ARMED','SECTION','LOCATION','ON LEAVE','LEAVE START DATE',
                'LEAVE END DATE', 'SPECIAL DUTY START DATE', 'SPECIAL DUTY END DATE']

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            font_style = xlwt.XFStyle()
            
            # for reports with specific Sections
            # query_parameter = "UN Women"
            # rows = Battallion_one.objects.filter(section=query_parameter).values_list('first_name', 'last_name', 'file_number', 'nin', 

            rows = Battallion_one.objects.filter().values_list('first_name', 'last_name', 'file_number', 'nin', 
                'ipps','account_number','contact','sex','rank','education_level','other_education_level','bank','branch','department',
                'title','status','shift','date_of_enlistment','date_of_transfer','date_of_promotion','date_of_birth','armed','section','location','on_leave','leave_start_date',
                'leave_end_date', 'special_duty_start_date', 'special_duty_end_date')

            for row in rows:
                row_num += 1

                for col_num in range(len(row)):
                    ws.write(row_num, col_num, str(row[col_num]), font_style)
            wb.save(response)

            return response
        else: 
            # print("You are not authorized to carry out this operation.")
            return JsonResponse({"detail": "You are not authorized to carry out this operation."}, status=401)
    except: 
        # print("You are not authorized to carry out this operation.")
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