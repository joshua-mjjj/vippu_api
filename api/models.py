import binascii
import os

from allauth.socialaccount.models import SocialAccount
from django.contrib.auth.base_user import BaseUserManager
from django.contrib.auth.models import AbstractUser
from django.core.exceptions import ValidationError
from django.db import models
from django.db.models import Avg
from model_utils.models import TimeStampedModel
from rest_framework_jwt.settings import api_settings

ACCOUNT_TYPES = (
    ("admin", "Admin"),
    ("in_charge", "In Charge"),
)

BATTALLION_TYPES = (
    ("battallion_one", "Battallion 1"),
    ("battallion_two", "Battallion 2"),
    ("battallion_three", "Battallion 3"),
    ("battallion_four", "Battallion 4"),
    ("battallion_five", "Battallion 5"),
    ("battallion_six", "Battallion 6"),
    ("all_battallions", "All Battallions"),
)

GENDER_TYPES = (
    ("male", "Male"),
    ("female", "Female"),
)

RANK_TYPES = (
    ("AIGP", "AIGP"),
    ("SCP", "SCP"),
    ("CP", "CP"),
    ("ACP", "ACP"),
    ("SSP", "SSP"),
    ("SP", "SP"),
    ("ASP", "ASP"),
    ("IP", "IP"),
    ("AIP", "AIP"),
    ("SGT", "SGT"),
    ("CPL", "CPL"),
    ("PC", "PC"),
    ("SPC", "SPC"),
)

BATTALLION_TWO_DEPARTMENT_TYPES = (
    ("embassy", "Embassy"),
    ("consolate", "Consolate"),
    ("high_commission", "High commission"),
    ("other_diplomats", "Other diplomats"),
    ("administration", "Administration")
)

TITLE_TYPES = (
    ("commandant", "Commandant"),
    ("deputy_commandant", "Deputy commandant"),
    ("staff_officer", "Staff officer"),
    ("head_of_operations", "Head of operations"),
    ("head_of_armoury", "Head of armoury"),
    ("supervisor", "Supervisor"),
    ("in_Charge", "In Charge"),
    ("2nd_in_charge", "2nd In Charge"),
    ("driver", "Driver"),
)

STATUS_TYPES = (
    ("active", "Active"),
    ("absent", "Absent(AWOL)"),
    ("transfered", "Transfered"),
    ("sick", "Sick"),
    ("dead", "Dead"),
    ("suspended", "Suspended"),
    ("dismissed", "Dismissed"),
    ("in_court", "In court"),
    ("deserted", "Deserted"),
    ("on_course", "On course"),
    ("on_mission", "On mission"),
)

SHIFT_TYPES = (
    ("day", "Day"),
    ("night", "Night"),
    ("long_night", "Long night"),
    ("none", "None(not applicable)"),
)

ARMED_TYPES = (
    ("yes", "Yes"),
    ("no", "No"),
)

LEAVE_TYPES = (
    ("pass_leave", "Pass leave"),
    ("maternity_leave", "Maternity leave"),
    ("sick_leave", "Sick leave"),
    ("study_leave", "Study leave"),
    ("annual_leave", "Annual leave"),
    ("not_no_leave", "Not no leave"),
)

EDUCATION_TYPES = (
    ("ple", "PLE"),
    ("uce", "UCE"),
    ("uace", "UACE"),
    ("diploma", "DIPLOMA"),
    ("bachelors", "BACHELORS"),
    ("masters", "MASTERS"),
    ("doctorate", "DOCTORATE(PhD)"),
    ("other", "OTHER")
)

GENDER_CHOICES = (("male", "Male"), ("female", "Female"))


class UserManager(BaseUserManager):
    def _create_user(self, username, password, is_staff, is_superuser, **extra_fields):
        if not username:
            raise ValueError("The given username must be set")
        # email = self.normalize_email(email)
        is_active = extra_fields.pop("is_active", True)
        user = self.model(
            username=username,
            is_staff=is_staff,
            is_active=is_active,
            is_superuser=is_superuser,
            **extra_fields
        )
        user.set_password(password)
        user.save(using=self._db)
        return user

    def create_user(self, username, password=None, **extra_fields):
        is_staff = extra_fields.pop("is_staff", False)
        return self._create_user(username, password, is_staff, False, **extra_fields)

    def create_superuser(self, username, password, **extra_fields):
        return self._create_user(username, password, True, True, **extra_fields)


class User(AbstractUser):
    username = models.CharField(max_length=100, unique=True) # db_index=True
    email = models.EmailField(
        "email address", max_length=255, unique=False, blank=True, null=True
    )
    first_name = models.CharField(max_length=32, blank=True, null=True)
    last_name = models.CharField(max_length=32, blank=True, null=True)

    account_type = models.CharField(max_length=32, choices=ACCOUNT_TYPES)
    battallion = models.CharField(max_length=32, choices=BATTALLION_TYPES)
    top_level_incharge = models.BooleanField(default=False) # Has access to the entire Battallion
    lower_level_incharge = models.BooleanField(default=False) # Has access to either Department or section in the Battallion they belong to
    department = models.CharField(max_length=32, blank=True, null=True) # Very Long choise field
    section = models.CharField(max_length=32, blank=True, null=True) # Very Long choise field

    phone_number = models.CharField(max_length=50, blank=True) # null=True
    is_staff = models.BooleanField("staff status", default=False)
    is_active = models.BooleanField("active", default=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    objects = UserManager()

    USERNAME_FIELD = "username" # Making username the required field
    REQUIRED_FIELDS = []

    @property
    def token(self):
        return self._generate_jwt_token()

    def _generate_jwt_token(self):
        jwt_payload_handler = api_settings.JWT_PAYLOAD_HANDLER
        jwt_encode_handler = api_settings.JWT_ENCODE_HANDLER

        payload = jwt_payload_handler(self)
        token = jwt_encode_handler(payload)

        return token

def generate_password_reset_code():
    return binascii.hexlify(os.urandom(20)).decode("utf-8")


# background_check_status = models.CharField(
#     max_length=20,
#     choices=[
#         ("pending", "Pending"),
#         ("started", "Started"),
#         ("passed", "Passed"),
#         ("failed", "Failed"),
#     ],
#     default="pending",
# )

# BATTALLION 2 TABLE MODEL 
class Battallion_two(models.Model):
    first_name = models.CharField(max_length=32)
    last_name = models.CharField(max_length=32)
    nin = models.CharField(max_length=32)
    ipps = models.CharField(max_length=32)
    file_number = models.CharField(max_length=32, unique=True) #, blank=True, null=True
    battallion = models.CharField(max_length=32)
    account_number = models.CharField(max_length=32, blank=True, null=True)
    contact = models.CharField(max_length=32, blank=True, null=True)
    sex = models.CharField(max_length=32, choices=GENDER_TYPES)
    rank = models.CharField(max_length=32, choices=RANK_TYPES)
    education_level = models.CharField(max_length=32, choices=EDUCATION_TYPES)
    other_education_level = models.CharField(max_length=32, blank=True, null=True) # Gives us an extra field incase of OTHER
    bank = models.CharField(max_length=32, blank=True, null=True)
    branch = models.CharField(max_length=32, blank=True, null=True)
    department = models.CharField(max_length=32, choices=BATTALLION_TWO_DEPARTMENT_TYPES)
    title = models.CharField(max_length=32, choices=TITLE_TYPES)
    status = models.CharField(max_length=32, choices=STATUS_TYPES)
    shift = models.CharField(max_length=32, choices=SHIFT_TYPES)
    date_of_enlistment = models.DateField(blank=True, null=True)
    date_of_transfer = models.DateField(blank=True, null=True)
    date_of_promotion = models.DateField(blank=True, null=True)
    date_of_birth = models.DateField(blank=False, null=False)
    armed = models.CharField(max_length=32, choices=ARMED_TYPES)
    section = models.CharField(max_length=32)
    location = models.CharField(max_length=32)
    on_leave = models.CharField(max_length=32, choices=LEAVE_TYPES)
    leave_start_date = models.DateField(blank=True, null=True) # Gives us an extra field 
    leave_end_date = models.DateField(blank=True, null=True) # Gives us an extra field 
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Battallion Two"
        verbose_name_plural = "Battallion Two"

    def __str__(self):
        return '{}, {}'.format(self.first_name, self.file_number)