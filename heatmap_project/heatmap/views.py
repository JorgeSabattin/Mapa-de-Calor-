import io
from django.shortcuts import render
from django.http import HttpResponse, JsonResponse
from django.views.decorators.csrf import csrf_exempt

from .procesador import procesar_archivo


def home(request):
    return render(request, 'heatmap/home.html')


def procesar(request):
    if request.method != 'POST':
        return JsonResponse({'error': 'Método no permitido'}, status=405)

    archivo = request.FILES.get('archivo')
    if not archivo:
        return render(request, 'heatmap/home.html', {'error': 'No se recibió ningún archivo.'})

    if not archivo.name.endswith('.xlsx'):
        return render(request, 'heatmap/home.html', {'error': 'Solo se aceptan archivos .xlsx'})

    try:
        resultado = procesar_archivo(archivo)

        # Guardar Excel en sesión como bytes (base64 para session)
        import base64
        excel_b64 = base64.b64encode(resultado['excel_bytes'].read()).decode('utf-8')
        request.session['excel_data'] = excel_b64
        request.session['excel_nombre'] = archivo.name.replace('.xlsx', '_heatmap.xlsx')

        context = {
            'heatmap_data':      resultado['heatmap_data'],
            'salas_libres_data': resultado['salas_libres_data'],
            'indice_data':       resultado['indice_data'],
            'stats':             resultado['stats'],
            'days':              resultado['days'],
            'nombre_archivo':    archivo.name,
        }
        return render(request, 'heatmap/resultados.html', context)

    except ValueError as e:
        return render(request, 'heatmap/home.html', {'error': str(e)})
    except Exception as e:
        return render(request, 'heatmap/home.html', {'error': f'Error al procesar: {str(e)}'})


def descargar_excel(request):
    import base64
    excel_b64 = request.session.get('excel_data')
    nombre    = request.session.get('excel_nombre', 'heatmap.xlsx')

    if not excel_b64:
        return HttpResponse('No hay archivo disponible. Procesa un archivo primero.', status=404)

    excel_bytes = base64.b64decode(excel_b64)
    response = HttpResponse(
        excel_bytes,
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = f'attachment; filename="{nombre}"'
    return response
