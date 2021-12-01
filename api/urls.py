import notifications.urls
from django.urls import include, path
from rest_framework.routers import DefaultRouter
from rest_framework_jwt.views import refresh_jwt_token
from rest_framework_swagger.views import get_swagger_view

from api.views import (
    AccountLoginAPIView,
    SignUp,
    UserProfile,
    ChangePasswordApi,
    UserType,
    BattallionTwoViewset,
    BattallionOneViewset,
    export_excel,
    BattalionTwo_query,
    BattalionOne_query,
    BattalionTwo_overrall,
    BattalionOne_overrall,
    BattalionTwo_section_query
)

router = DefaultRouter()
router.register(r"battallion_two", BattallionTwoViewset, basename="battallion_two")
router.register(r"battallion_one", BattallionOneViewset, basename="battallion_one")

schema_view = get_swagger_view(title="VIPPU API")

urlpatterns = [
    path("login/", AccountLoginAPIView.as_view(), name="login"),
    path("signup/", SignUp.as_view(), name="api-signup"),
    path("user_type_check/", UserType.as_view(), name="user-check"),
    path("battalionquery_two/", BattalionTwo_query.as_view(), name="battalionquery_two"),
    path("battalionquery_one/", BattalionOne_query.as_view(), name="battalionquery_one"),
    path("battaliontwo_overrall/", BattalionTwo_overrall.as_view(), name="battaliontwo_overrall"),
    path("battalionone_overrall/", BattalionOne_overrall.as_view(), name="battalionone_overrall"),
    path("battalionone_section_query/", BattalionTwo_section_query.as_view(), name="battalionone_section_query"),
    path("export_excel/", export_excel, name="export_excel"),

    path("token/refresh/", refresh_jwt_token),
    path("users/me/", UserProfile.as_view()),
    path("auth/", include("rest_framework.urls", namespace="rest_framework")),
    path("auth/", include("rest_framework_social_oauth2.urls")),
    path("auth/password/change/", ChangePasswordApi.as_view()),
    path("", include(router.urls)),
]


# path("auth/accounts/confirm/", AccountActivation.as_view()),
# path("auth/password/request/reset/", RequestPasswordReset.as_view()),
# path("auth/password/reset/request/confirm/", ConfirmPasswordResetRequest.as_view()),
# path("auth/password/reset/", ResetPasswordApi.as_view()),