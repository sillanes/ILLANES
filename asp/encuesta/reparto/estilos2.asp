<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css" />
    
    <style>
        * { box-sizing: border-box; font-family: Arial, sans-serif; }
        body { margin: 0; background-color: #f4f4f4; }
        
        .sidebar {
            width: 220px;
            background-color: #2c3e50;
            color: white;
            position: fixed;
            height: 100%;
            top: 0;
            left: 0;
            padding-top: 60px; /* espacio para header fijo */
            transition: left 0.3s ease;
            overflow-y: auto;
            z-index: 1001;
        }
        .sidebar h2 {
            text-align: center;
            margin-bottom: 30px;
        }
        .sidebar a {
            display: block;
            color: white;
            padding: 12px 20px;
            text-decoration: none;
        }
        .sidebar a:hover {
            background-color: #34495e;
        }
        .sidebar i {
            margin-right: 10px;
        }

        header {
            background-color: #2c3e50;
            color: white;
            padding: 10px 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            position: fixed;
            width: 100%;
            top: 0;
            left: 0;
            height: 60px;
            z-index: 1100;
        }
        .menu-toggle {
            display: none;
            background: none;
            border: none;
            color: white;
            font-size: 22px;
            cursor: pointer;
        }
        .logout {
            background-color: #e74c3c;
            color: white;
            border: none;
            padding: 8px 14px;
            cursor: pointer;
            border-radius: 4px;
            font-size: 14px;
        }

        .main-content {
            margin-left: 220px;
            padding: 80px 20px 40px 20px; /* padding top para header fijo */
            transition: margin-left 0.3s ease;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            background: white;
            margin-top: 20px;
        }
        th, td {
            border: 1px solid #ccc;
            padding: 10px;
            text-align: left;
        }
        th {
            background-color: #ecf0f1;
        }
        .icon-btn {
            background: none;
            border: none;
            cursor: pointer;
            font-size: 18px;
            margin-right: 8px;
            color: inherit;
            text-decoration: none;
        }
        .icon-btn.entregar { color: green; }
        .icon-btn.rechazar { color: red; }

        @media (max-width: 768px) {
            .sidebar {
                left: -220px; /* oculto por defecto */
            }
            .sidebar.open {
                left: 0;
            }
            .menu-toggle {
                display: inline;
            }
            .main-content {
                margin-left: 0 !important;
                padding: 80px 15px 40px 15px;
                overflow-x: auto; /* scroll horizontal en móvil si tabla ancha */
            }
            table {
                font-size: 14px;
                min-width: 600px; /* fuerza scroll horizontal si hace falta */
            }
        }
		
/* Contenedor con scroll horizontal para tabla */
.table-responsive {
    width: 100%;
    overflow-x: auto;
    -webkit-overflow-scrolling: touch; /* para smooth scrolling en iOS */
}

/* Estilo base tabla (ya tienes) */
table {
    width: 100%;
    border-collapse: collapse;
    background: white;
    margin-top: 20px;
    min-width: 600px; /* mínimo para que las columnas no queden muy chicas */
}

/* Ajustes móviles */
@media (max-width: 768px) {
    table {
        font-size: 14px;
        min-width: 0; /* para permitir que se reduzca */
    }
    th, td {
        padding: 8px 6px;
    }
}
		
    </style>