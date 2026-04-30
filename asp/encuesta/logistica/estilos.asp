<style>
body {
    margin: 0;
    font-family: Arial, sans-serif;
    background: #f4f4f4;
}

header {
    background-color: #1e90ff;
    color: white;
    padding: 10px;
    display: flex;
    align-items: center;
    justify-content: space-between;
}

.menu-toggle {
    background: none;
    border: none;
    font-size: 1.5rem;
    color: white;
    cursor: pointer;
}

.logout {
    background-color: #ff4d4d;
    color: white;
    border: none;
    padding: 5px 10px;
    cursor: pointer;
}

.sidebar {
    background-color: #333;
    color: white;
    width: 220px;
    height: 100vh;
    position: fixed;
    left: 0;
    top: 0;
    padding-top: 60px;
    transition: transform 0.3s ease;
}

.sidebar ul {
    list-style-type: none;
    padding: 0;
}

.sidebar li {
    padding: 15px;
}

.sidebar li a {
    color: white;
    text-decoration: none;
    display: block;
}

.sidebar li a:hover {
    background-color: #444;
}

.sidebar.open {
    transform: translateX(0);
}

.main-content {
    margin-left: 220px;
    padding: 20px;
}

.table-responsive {
    overflow-x: auto;
    background: white;
    padding: 10px;
    border-radius: 8px;
}

table {
    width: 100%;
    border-collapse: collapse;
}

table th, table td {
    padding: 10px;
    border: 1px solid #ddd;
    text-align: center;
}

.action-icon {
    margin: 0 5px;
    cursor: pointer;
    color: #333;
}

.action-icon.trabajar:hover {
    color: #007bff;
}

.action-icon.cerrar:hover {
    color: #28a745;
}

/* Responsive */
@media (max-width: 768px) {
    .sidebar {
        transform: translateX(-100%);
        position: fixed;
        z-index: 1000;
    }

    .sidebar.open {
        transform: translateX(0);
    }

    .main-content {
        margin-left: 0;
    }

    header {
        flex-direction: column;
        align-items: flex-start;
    }
}
</style>
