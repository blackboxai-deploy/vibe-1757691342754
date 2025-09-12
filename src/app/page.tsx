"use client";

import { useState } from 'react';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Separator } from '@/components/ui/separator';
import { Badge } from '@/components/ui/badge';
import { createExcelFile, downloadExcelFile, type ExcelData } from '@/lib/excelUtils';

export default function Home() {
  const [formData, setFormData] = useState<ExcelData>({
    organizacion: 'Sociedad de Beneficencia de Puno',
    mes: 'Enero',
    año: new Date().getFullYear().toString(),
    fechaElaboracion: (() => {
      const today = new Date();
      const day = today.getDate().toString().padStart(2, '0');
      const month = (today.getMonth() + 1).toString().padStart(2, '0');
      const year = today.getFullYear();
      return `${day}/${month}/${year}`;
    })(),
    saldoInicialMes: 0,
    fondoRotatorioFijo: 0
  });

  const [isGenerating, setIsGenerating] = useState(false);

  const handleInputChange = (field: keyof ExcelData, value: string | number) => {
    setFormData(prev => ({
      ...prev,
      [field]: value
    }));
  };

  const handleGenerateExcel = async () => {
    setIsGenerating(true);
    try {
      const workbook = createExcelFile(formData);
      const filename = `Anexo_Practico_${formData.mes}_${formData.año}.xlsx`;
      downloadExcelFile(workbook, filename);
    } catch (error) {
      console.error('Error generando archivo Excel:', error);
      alert('Error al generar el archivo Excel. Por favor intente nuevamente.');
    } finally {
      setIsGenerating(false);
    }
  };

  const meses = [
    'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
    'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'
  ];

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-gray-100">
      {/* Header */}
      <header className="bg-slate-800 text-white py-6">
        <div className="container mx-auto px-4">
          <div className="flex items-center justify-between">
            <div>
              <h1 className="text-2xl font-bold">Sociedad de Beneficencia de Puno</h1>
              <p className="text-slate-300 mt-1">Sistema de Generación de Anexos Prácticos</p>
            </div>
            <div className="text-right">
              <Badge variant="secondary" className="bg-blue-600 text-white">
                Reglamento Interno
              </Badge>
            </div>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="container mx-auto px-4 py-8 max-w-4xl">
        <div className="grid gap-8">
          
          {/* Información del Sistema */}
          <Card>
            <CardHeader>
              <CardTitle className="text-slate-800">Generador de Anexo Práctico</CardTitle>
              <CardDescription>
                Sistema automatizado para crear archivos Excel con los formatos establecidos 
                en el Reglamento Interno de Negocios
              </CardDescription>
            </CardHeader>
            <CardContent>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="space-y-2">
                  <h4 className="font-semibold text-slate-700">Hojas incluidas:</h4>
                  <ul className="text-sm text-slate-600 space-y-1">
                    <li>• Portada institucional</li>
                    <li>• Rendición Mensual</li>
                    <li>• Flujo de Caja Mensual</li>
                  </ul>
                </div>
                <div className="space-y-2">
                  <h4 className="font-semibold text-slate-700">Características:</h4>
                  <ul className="text-sm text-slate-600 space-y-1">
                    <li>• Registro de Compras</li>
                    <li>• Inventario automatizado</li>
                    <li>• Solicitud de Reposición</li>
                  </ul>
                </div>
              </div>
            </CardContent>
          </Card>

          {/* Formulario de Datos */}
          <Card>
            <CardHeader>
              <CardTitle className="text-slate-800">Configuración del Reporte</CardTitle>
              <CardDescription>
                Complete los datos iniciales para generar el archivo Excel personalizado
              </CardDescription>
            </CardHeader>
            <CardContent className="space-y-6">
              
              {/* Información Básica */}
              <div className="space-y-4">
                <h4 className="font-semibold text-slate-700 border-b pb-2">Información del Período</h4>
                
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div className="space-y-2">
                    <Label htmlFor="mes">Mes del reporte</Label>
                    <select
                      id="mes"
                      className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                      value={formData.mes}
                      onChange={(e) => handleInputChange('mes', e.target.value)}
                    >
                      {meses.map(mes => (
                        <option key={mes} value={mes}>{mes}</option>
                      ))}
                    </select>
                  </div>
                  
                  <div className="space-y-2">
                    <Label htmlFor="año">Año</Label>
                    <Input
                      id="año"
                      type="number"
                      value={formData.año}
                      onChange={(e) => handleInputChange('año', e.target.value)}
                      className="focus:ring-2 focus:ring-blue-500"
                      min="2020"
                      max="2030"
                    />
                  </div>
                </div>

                <div className="space-y-2">
                  <Label htmlFor="fechaElaboracion">Fecha de elaboración</Label>
                  <Input
                    id="fechaElaboracion"
                    type="date"
                    value={(() => {
                      const fechaParts = formData.fechaElaboracion.split('/');
                      if (fechaParts.length === 3) {
                        // Formato: DD/MM/YYYY -> YYYY-MM-DD
                        const [day, month, year] = fechaParts;
                        return `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`;
                      }
                      return new Date().toISOString().split('T')[0];
                    })()}
                    onChange={(e) => {
                      const date = new Date(e.target.value);
                      const day = date.getDate().toString().padStart(2, '0');
                      const month = (date.getMonth() + 1).toString().padStart(2, '0');
                      const year = date.getFullYear();
                      handleInputChange('fechaElaboracion', `${day}/${month}/${year}`);
                    }}
                    className="focus:ring-2 focus:ring-blue-500"
                  />
                  <p className="text-xs text-slate-500">
                    Fecha actual: {formData.fechaElaboracion}
                  </p>
                </div>
              </div>

              <Separator />

              {/* Datos Financieros */}
              <div className="space-y-4">
                <h4 className="font-semibold text-slate-700 border-b pb-2">Configuración Financiera</h4>
                
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div className="space-y-2">
                    <Label htmlFor="saldoInicial">Saldo inicial del mes (S/.)</Label>
                    <Input
                      id="saldoInicial"
                      type="number"
                      step="0.01"
                      value={formData.saldoInicialMes}
                      onChange={(e) => handleInputChange('saldoInicialMes', parseFloat(e.target.value) || 0)}
                      className="focus:ring-2 focus:ring-blue-500"
                      placeholder="0.00"
                    />
                    <p className="text-xs text-slate-500">
                      Monto con el que inicia el período contable
                    </p>
                  </div>
                  
                  <div className="space-y-2">
                    <Label htmlFor="fondoRotatorio">Fondo rotatorio fijo (S/.)</Label>
                    <Input
                      id="fondoRotatorio"
                      type="number"
                      step="0.01"
                      value={formData.fondoRotatorioFijo}
                      onChange={(e) => handleInputChange('fondoRotatorioFijo', parseFloat(e.target.value) || 0)}
                      className="focus:ring-2 focus:ring-blue-500"
                      placeholder="0.00"
                    />
                    <p className="text-xs text-slate-500">
                      Monto fijo que debe mantenerse en el fondo
                    </p>
                  </div>
                </div>
              </div>

              <Separator />

              {/* Información Adicional */}
              <div className="bg-blue-50 p-4 rounded-lg">
                <h4 className="font-semibold text-slate-800 mb-2">Características del Archivo</h4>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-sm text-slate-600">
                  <div>
                    <p className="font-medium text-slate-700">Fórmulas automáticas:</p>
                    <ul className="list-disc list-inside space-y-1">
                      <li>Cálculo de saldos acumulativos</li>
                      <li>Totales de ingresos y egresos</li>
                      <li>Utilidad neta automática</li>
                    </ul>
                  </div>
                  <div>
                    <p className="font-medium text-slate-700">Validaciones incluidas:</p>
                    <ul className="list-disc list-inside space-y-1">
                      <li>Listas desplegables</li>
                      <li>Protección de fórmulas</li>
                      <li>Formato institucional</li>
                    </ul>
                  </div>
                </div>
              </div>

              {/* Botón de Generación */}
              <div className="flex justify-center pt-4">
                <Button
                  onClick={handleGenerateExcel}
                  disabled={isGenerating}
                  className="px-8 py-3 bg-slate-800 hover:bg-slate-700 text-white font-semibold text-lg"
                  size="lg"
                >
                  {isGenerating ? (
                    <>
                      <div className="animate-spin rounded-full h-4 w-4 border-b-2 border-white mr-2"></div>
                      Generando archivo...
                    </>
                  ) : (
                    'Generar Archivo Excel'
                  )}
                </Button>
              </div>

              <p className="text-xs text-center text-slate-500 mt-4">
                El archivo se descargará automáticamente con el nombre: 
                Anexo_Practico_{formData.mes}_{formData.año}.xlsx
              </p>
            </CardContent>
          </Card>

          {/* Instrucciones */}
          <Card>
            <CardHeader>
              <CardTitle className="text-slate-800">Instrucciones de Uso</CardTitle>
            </CardHeader>
            <CardContent>
              <div className="space-y-4 text-sm text-slate-600">
                <div>
                  <h4 className="font-semibold text-slate-700 mb-2">Después de descargar:</h4>
                  <ol className="list-decimal list-inside space-y-1">
                    <li>Abra el archivo Excel generado</li>
                    <li>Complete los datos en las hojas según sus registros</li>
                    <li>Las fórmulas calcularán automáticamente los totales</li>
                    <li>No elimine las celdas con fórmulas (están protegidas)</li>
                    <li>Use las listas desplegables para tipos de compra</li>
                  </ol>
                </div>

                <div className="bg-blue-50 p-3 rounded border-l-4 border-blue-400">
                  <p className="font-medium text-blue-800">Fórmulas entre hojas (Instrucciones específicas):</p>
                  <div className="text-blue-700 space-y-2 mt-2">
                    <p><strong>En Flujo de Caja Mensual:</strong></p>
                    <ul className="list-disc list-inside ml-4 space-y-1">
                      <li>Celda B5: Escriba <code>=SUM('Rendición Mensual'.E:E)</code></li>
                      <li>Celda B6: Escriba <code>=SUM('Rendición Mensual'.F:F)</code></li>
                    </ul>
                    <p><strong>Para Solicitud de Reposición:</strong></p>
                    <ul className="list-disc list-inside ml-4">
                      <li>Complete manualmente los montos según sus registros</li>
                    </ul>
                  </div>
                </div>
                
                <div className="bg-amber-50 p-3 rounded border-l-4 border-amber-400">
                  <p className="font-medium text-amber-800">Importante:</p>
                  <p className="text-amber-700">
                    Mantenga una copia de seguridad del archivo original antes de realizar modificaciones.
                    Las instrucciones detalladas aparecen en cada hoja del Excel generado.
                  </p>
                </div>

                <div className="bg-green-50 p-3 rounded border-l-4 border-green-400">
                  <p className="font-medium text-green-800">✅ Correcciones aplicadas:</p>
                  <ul className="text-green-700 list-disc list-inside space-y-1 mt-1">
                    <li>Visualización de fecha corregida</li>
                    <li>Fórmulas problemáticas reemplazadas con valores iniciales</li>
                    <li>Instrucciones incluidas en cada hoja del Excel</li>
                    <li>Referencias entre hojas simplificadas</li>
                  </ul>
                </div>
              </div>
            </CardContent>
          </Card>

        </div>
      </main>

      {/* Footer */}
      <footer className="bg-slate-800 text-white py-4 mt-12">
        <div className="container mx-auto px-4 text-center">
          <p className="text-slate-300">
            © 2024 Sociedad de Beneficencia de Puno - Sistema de Gestión Financiera
          </p>
        </div>
      </footer>
    </div>
  );
}