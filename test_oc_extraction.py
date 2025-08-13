#!/usr/bin/env python3
"""
Script de prueba para verificar la extracción de números de OC
"""

import re

def extraer_numero_oc(filename):
    """
    Extrae el número de OC del inicio del nombre del archivo.
    
    Formato esperado: "NúmeroOC NúmeroEntrega"
    Ejemplos:
    - "5600025440 0082214777" -> "5600025440"
    - "5100064001 0082214479" -> "5100064001"
    - "5600025443 0082214783" -> "5600025443"
    
    Args:
        filename (str): Nombre del archivo sin extensión
        
    Returns:
        str: Número de OC extraído o None si no se puede extraer
    """
    try:
        # Buscar el primer número al inicio del archivo (antes del espacio)
        match = re.match(r'^(\d+)', filename)
        if match:
            oc_number = match.group(1)
            # Verificar que sea un número razonable (entre 8 y 12 dígitos)
            if 8 <= len(oc_number) <= 12:
                return oc_number
        
        # Si no encuentra al inicio, buscar el primer número en el archivo
        matches = re.findall(r'\b(\d{8,12})\b', filename)
        if matches:
            return matches[0]  # Retornar el primer número encontrado
        
        return None
        
    except Exception as e:
        print(f"Error extrayendo OC de {filename}: {e}")
        return None

def test_oc_extraction():
    """Prueba la función de extracción de OC"""
    
    # Casos de prueba
    test_cases = [
        "5600025440 0082214777",
        "5600025443 0082214783", 
        "5100064001 0082214479",
        "5100064002 0082214653",
        "5600025441 0082214778",
        "5600025442 0082214782",
        "OC5100064001 0082214479",  # Formato antiguo
        "test_file.xlsx",  # Caso inválido
        "12345678 99999999",  # Número muy corto
        "123456789012345 99999999",  # Número muy largo
    ]
    
    print("=== PRUEBA DE EXTRACCIÓN DE OC ===")
    print(f"{'Archivo':<30} {'OC Extraída':<15} {'Resultado'}")
    print("-" * 60)
    
    for test_file in test_cases:
        oc = extraer_numero_oc(test_file)
        if oc:
            print(f"{test_file:<30} {oc:<15} ✅ VÁLIDO")
        else:
            print(f"{test_file:<30} {'None':<15} ❌ INVÁLIDO")
    
    print("-" * 60)
    print("✅ Prueba completada")

if __name__ == "__main__":
    test_oc_extraction() 