# -*- coding: utf-8 -*-
import io
from rest_framework.response import Response
from rest_framework import views, status
from django.utils.decorators import method_decorator
from django.views.decorators.csrf import csrf_exempt
from django.views import generic


class GenericModelExport(views.APIView):

    @classmethod
    def as_view(cls, **initkwargs):
        from django.views.decorators.csrf import csrf_exempt
        """
        Store the original class on the view function.

        This allows us to discover information about the view when we do URL
        reverse lookups.  Used for breadcrumb generation.
        """

        view = super(views.APIView, cls).as_view(**initkwargs)
        view.cls = cls
        view.initkwargs = initkwargs

        # Note: session based authentication is explicitly CSRF validated,
        # all other authentication is CSRF exempt.
        return csrf_exempt(view)

    @method_decorator(csrf_exempt)
    def dispatch(self, request, *args, **kwargs):
        return views.APIView.dispatch(self, request, *args, **kwargs)

    def get_data_config(self, **kwargs):
        if hasattr(self, "data_config"):
            return self.data_config

        queryset = self.get_queryset()
        model = queryset.model
        fields = getattr(self, "fields", "__all__")

        if fields == "__all__":
            print("aun no implementado")
            return None
        elif type(fields) in [list, tuple]:
            data_config = []
            for field in fields:
                try:
                    verbose_name = model._meta.get_field(field).verbose_name
                except Exception:
                    verbose_name = field
                data_config.append([verbose_name, field])
            return data_config
        return None

    def get_queryset(self, **kwargs):
        if hasattr(self, "queryset"):
            queryset = self.queryset
            return queryset

    def get_data(self, request, **kwargs):
        header_format = getattr(self, "header_format", {})

        data_config = self.get_data_config()

        if not data_config:
            return Response(status=status.HTTP_204_NO_CONTENT)

        headers = [{"text": config[0], "format": header_format}
                   for config in data_config]
        obj_attrs = [config[1] for config in data_config]

        columns_width_pixel = getattr(self, "columns_width_pixel", None)
        columns_width = getattr(self, "columns_width", None)
        if isinstance(columns_width_pixel, bool) and columns_width_pixel:
            try:
                self.columns_width_pixel = [config[2]
                                            for config in data_config]
            except Exception:
                pass
        if isinstance(columns_width, bool) and columns_width:
            try:
                self.columns_width = [config[2] for config in data_config]
            except Exception:
                pass

        data = [headers]

        queryset = list(self.get_queryset(**kwargs))

        for obj in queryset:
            obj_data = []
            for obj_attr in obj_attrs:
                obj_data.append(get_attr(obj, obj_attr, **kwargs))
            data.append(obj_data)
        return data

    def get_file_name(self, request, **kwargs):
        return getattr(self, "xlsx_name", "export")

    def get(self, request, **kwargs):
        from django.template.defaultfilters import slugify
        self.request = request

        data = self.get_data(request, **kwargs)

        if request.query_params.get("test") == "true":
            return Response(data)

        file_name = self.get_file_name(request, **kwargs)
        slug_file_name = slugify(file_name)
        tab_name = getattr(self, "tab_name", "list")
        columns_width = getattr(self, "columns_width", [])
        columns_width_pixel = getattr(self, "columns_width_pixel", [])
        export_xlsx(name="%s.xlsx" % slug_file_name,
                    data=[{"name": tab_name, "table_data": data,
                           "columns_width": columns_width,
                           "columns_width_pixel": columns_width_pixel}])

        from wsgiref.util import FileWrapper
        from django.http import HttpResponse
        try:
            file_xlsx = open("%s.xlsx" % slug_file_name, 'rb')
        except Exception as e:
            return Response({"errors": [u"%s" % e]},
                            status=status.HTTP_400_BAD_REQUEST)
        response = HttpResponse(FileWrapper(file_xlsx),
                                content_type='application/vnd.ms-excel')
        response['Content-Disposition'] = 'attachment; filename="%s"' % (
            "%s.xlsx" % file_name)
        return response


class GenericBasicExport(views.APIView):

    @classmethod
    def as_view(cls, **initkwargs):
        from django.views.decorators.csrf import csrf_exempt
        """
        Store the original class on the view function.

        This allows us to discover information about the view when we do URL
        reverse lookups.  Used for breadcrumb generation.
        """

        view = super(views.APIView, cls).as_view(**initkwargs)
        view.cls = cls
        view.initkwargs = initkwargs

        # Note: session based authentication is explicitly CSRF validated,
        # all other authentication is CSRF exempt.
        return csrf_exempt(view)

    @method_decorator(csrf_exempt)
    def dispatch(self, request, *args, **kwargs):
        return generic.View.dispatch(self, request, *args, **kwargs)

    def get_data(self, request, **kwargs):
        if hasattr(self, "data"):
            data = self.data
            return data
        return None

    def get_file_name(self, request, **kwargs):
        return getattr(self, "xlsx_name", "export")

    def get(self, request, **kwargs):
        from django.template.defaultfilters import slugify
        self.request = request
        data = self.get_data(request, **kwargs)
        if request.query_params.get("test") == "true":
            return Response(data)
        if not data:
            return Response({"system": "sin datos configurados"})

        file_name = self.get_file_name(request, **kwargs)
        slug_file_name = slugify(file_name)
        export_xlsx(name="%s.xlsx" % slug_file_name, data=data)

        from wsgiref.util import FileWrapper
        from django.http import HttpResponse
        try:
            file_xlsx = open("%s.xlsx" % slug_file_name, 'rb')
        except Exception as e:
            return Response({"errors": [u"%s" % e]},
                            status=status.HTTP_400_BAD_REQUEST)
        response = HttpResponse(FileWrapper(file_xlsx),
                                content_type='application/vnd.ms-excel')
        response['Content-Disposition'] = 'attachment; filename="%s"' % (
            "%s.xlsx" % file_name)
        return response

    def get_data_model_class(self, request, model_class, **kwargs):
        if not model_class:
            return {}
        model_class.request = request
        tab_name = getattr(model_class, "tab_name", "list")
        data = model_class.get_data(request, **kwargs)
        columns_width = getattr(model_class, "columns_width", [])
        columns_width_pixel = getattr(
            model_class, "columns_width_pixel", [])
        return {"name": tab_name, "table_data": data,
                "columns_width": columns_width,
                "columns_width_pixel": columns_width_pixel}


class FastModelExport(GenericModelExport):

    def get(self, request, **kwargs):
        from django.template.defaultfilters import slugify
        self.request = request

        file_name = self.get_file_name(request, **kwargs)
        slug_file_name = slugify("fast " + file_name)

        from wsgiref.util import FileWrapper
        from django.http import HttpResponse
        try:
            file_xlsx = open("%s.xlsx" % slug_file_name, 'rb')
        except Exception as e:
            return Response({"errors": [u"%s" % e]},
                            status=status.HTTP_400_BAD_REQUEST)
        response = HttpResponse(FileWrapper(file_xlsx),
                                content_type='application/vnd.ms-excel')
        response['Content-Disposition'] = 'attachment; filename="%s"' % (
            "%s.xlsx" % file_name)
        return response

    def post(self, request, **kwargs):
        from django.template.defaultfilters import slugify
        self.request = request

        data = self.get_data(request, **kwargs)

        if request.query_params.get("test") == "true":
            return Response(data)

        file_name = self.get_file_name(request, **kwargs)
        slug_file_name = slugify("fast " + file_name)
        tab_name = getattr(self, "tab_name", "list")
        columns_width = getattr(self, "columns_width", [])
        columns_width_pixel = getattr(self, "columns_width_pixel", [])
        export_xlsx(name="%s.xlsx" % slug_file_name,
                    data=[{"name": tab_name, "table_data": data,
                           "columns_width": columns_width,
                           "columns_width_pixel": columns_width_pixel}])

        return Response({"msg": "Se genero el archivo %s" % (file_name)})


class FastBasicExport(GenericBasicExport):

    def get(self, request, **kwargs):
        from django.template.defaultfilters import slugify
        self.request = request

        file_name = self.get_file_name(request, **kwargs)
        slug_file_name = slugify("fast " + file_name)

        from wsgiref.util import FileWrapper
        from django.http import HttpResponse
        try:
            file_xlsx = open("%s.xlsx" % slug_file_name, 'rb')
        except Exception as e:
            return Response({"errors": [u"%s" % e]},
                            status=status.HTTP_400_BAD_REQUEST)
        response = HttpResponse(FileWrapper(file_xlsx),
                                content_type='application/vnd.ms-excel')
        response['Content-Disposition'] = 'attachment; filename="%s"' % (
            "%s.xlsx" % file_name)
        return response

    def post(self, request, **kwargs):
        from django.template.defaultfilters import slugify
        self.request = request
        print("entro a fast")
        data = self.get_data(request, **kwargs)
        if request.query_params.get("test") == "true":
            return Response(data)
        if not data:
            return Response({"system": "sin datos configurados"})

        print("genero data")
        file_name = self.get_file_name(request, **kwargs)
        slug_file_name = slugify("fast " + file_name)
        export_xlsx(name="%s.xlsx" % slug_file_name, data=data)
        print("genero archivo")

        return Response({"msg": "Se genero el archivo %s" % (file_name)})


def export_xlsx(name="test.xlsx", data=None, in_memory=False):
    import xlsxwriter
    output = None
    if in_memory:
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {"in_memory": True})
    else:
        workbook = xlsxwriter.Workbook(name)

    id_index = 0
    if data is None:
        data = []
    for worksheet_data in data:
        id_index += 1
        name = worksheet_data.get("name", "page %s" % id_index)
        table_data = worksheet_data.get("table_data", [])
        columns_width = worksheet_data.get("columns_width")
        columns_width_pixel = worksheet_data.get("columns_width_pixel")
        max_decimal = worksheet_data.get("max_decimal", 1)
        # Create a workbook and add a worksheet.
        # ----------------------------------------------------------
        worksheet = workbook.add_worksheet(name)
        if columns_width:
            col = 0
            for column_width in columns_width:
                worksheet.set_column(col, col, column_width)
                col += 1
        elif columns_width_pixel:
            col = 0
            for column_width in columns_width_pixel:
                try:
                    worksheet.set_column(col, col, int(column_width / 7.5))
                except Exception as e:
                    print(e)
                    continue
                col += 1

        row = 0
        for linea in table_data:
            col = 0
            for celda in linea:
                if type(celda) in [str]:
                    if celda[0:1] == "=":
                        worksheet.write_formula(row, col, celda)
                        col += 1
                        continue
                elif type(celda) in [float]:
                    worksheet.write(row, col, float(("{:.%sf}" % max_decimal)
                                                    .format(celda)))
                    col += 1
                    continue
                elif isinstance(celda, dict):
                    text = celda.get("text", "")
                    format_config = celda.get("format")
                    if isinstance(format_config, dict):
                        cell_format = workbook.add_format(format_config)
                        worksheet.write(row, col, text, cell_format)
                    else:
                        worksheet.write(row, col, text)
                    col += 1
                    continue
                worksheet.write(row, col, celda)
                col += 1
            row += 1
    workbook.close()
    if in_memory:
        output.seek(0)
        return output


def get_attr(obj, attr, **kwargs):
    config = kwargs.pop("config", {})
    if type(attr) in [str]:
        attrs = attr.split(".")
    elif isinstance(attr, list):
        attrs = attr
    try:
        first_attr = attrs.pop(0)
    except Exception:
        return ""
    value = getattr(obj, first_attr, "")
    if attrs:
        if kwargs:
            kwargs["config"] = config
        else:
            kwargs = {"config": config}
        return get_attr(value, attrs, **kwargs)
    else:
        from datetime import datetime, date
        import decimal
        if type(value) in [str, int, float, decimal.Decimal]:
            if first_attr in config:
                return config[first_attr].get(value, value)
            return value

        elif isinstance(value, bool):
            if first_attr in config:
                return config[first_attr].get(value, value)
            return u"Sí" if value else u"No"

        elif type(value) in [date]:
            return value.strftime("%d/%m/%Y")
        elif type(value) in [datetime]:
            value = get_datetime_mx(value)
            return value.strftime("%d/%m/%Y %H:%M:%S")
        elif str(type(value)) == "<type 'instancemethod'>":
            try:
                return u"%s" % value(**kwargs)
            except Exception:
                try:
                    return u"%s" % value()
                except Exception as e:
                    print(e)
                    return ""
        elif not value:
            return ""
        else:
            print(type(value))
            return u"%s" % value


def get_datetime_mx(datetime_utc):
    import pytz
    cdmx_tz = pytz.timezone("America/Mexico_City")
    return datetime_utc.astimezone(cdmx_tz)
