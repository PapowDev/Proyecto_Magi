import sqlite3

class Database:
    def __init__(self, db_name='ventas.db'):
        self.conn = sqlite3.connect(db_name)
        self.cursor = self.conn.cursor()
        self._crear_tablas()

    def _crear_tablas(self):
        # Crear la tabla productos si no existe
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS productos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT NOT NULL UNIQUE,
            nombre TEXT NOT NULL,
            cantidad_por_caja INTEGER NOT NULL,
            prioridad TEXT DEFAULT 'baja'
        )
        ''')

        # Insertar datos de ejemplo en la tabla productos (solo si está vacía)
        self.cursor.execute('SELECT COUNT(*) FROM productos')
        if self.cursor.fetchone()[0] == 0:
            productos_ejemplo = [
                ('FB007', 'DXN Morinzhi', 36),
                ('FB027', 'DXN Morinzyme', 36),
                ('FB063', 'DXN Zhi Café Classic', 12),
                ('FB069', 'DXN Cordyceps Cereal', 12),
                ('FB097', 'DXN Lingzhi Tea Latte', 40),
                ('FB098', 'DXN White Coffee Zhino', 40),
                ('FB109', 'DXN Lingzhi Coffee 3 in 1', 40),
                ('FB121', 'DXN Spica Tea', 60),
                ('FB122', 'DXN Lingzhi Black Coffee', 80),
                ('FB124', 'DXN Cocozhi', 25),
                ('FB125', 'DXN Spirulina Cereal', 15),
                ('FB126', 'DXN Zhi Mocha', 40),
                ('FB128', 'DXN Vita Café', 40),
                ('FB143', 'DXN Zhi Mint Plus', 10),
                ('FB150', 'DXN Civattino Coffee', 36),
                ('FB155', 'DXN Lemonzhi', 40),
                ('FB223', 'DXN Zhitea', 24),
                ('FB229', 'DXN Oocha', 22),
                ('FB303', 'DXN VCO-L 285ml', 12),
                ('FB308', 'DXN Oozhi Tea', 24),
                ('FB351', 'DXN Oozhi Tea 30g', 40),
                ('FB352', 'DXN Oozhi Tea (30\'s)', 80),
                ('FB360', 'DXN Sugar', 40),
                ('FB361', 'DXN Non Dairy Creamer', 30),
                ('FB362', 'DXN Oolong Tea Powder 30g', 90),
                ('FB369', 'DXN Ootea Lingzhi Coffee Mix 3 in 1', 40),
                ('FB370', 'DXN Ootea Lingzhi Coffee Mix 3 in 1 Lite', 40),
                ('FB371', 'DXN Ootea Lingzhi Coffee Mix 2 in 1', 36),
                ('FB372', 'DXN Ootea Cordyceps Coffee Mix 3 in 1', 40),
                ('FB373', 'DXN Ootea Zhi Mocha Mix', 40),
                ('FB375', 'DXN Ootea Lingzhi Black Coffee Mix', 80),
                ('FB376', 'DXN Ootea Vita Cafe Mix', 40),
                ('FB377', 'DXN Ootea Eu Cafe Mix', 40),
                ('FB397', 'DXN Spirulina Coffee', 36),
                ('FB432', 'DXN Nutrizhi', 20),
                ('FB437', 'DXN Spirulina Buckwheat Noodle', 40),
                ('HF039', 'DXN MycoVeggie', 9),
                ('HF007', 'RG Powder', 60),
                ('HF008', 'GL Powder', 60),
                ('FB442', 'DXN Chinese Cherry Sparkling Drink', 24),
                ('FB443', 'DXN Apple Sparkling Drink', 24)
            ]
            
            # Usar executemany para insertar todos los productos
            self.cursor.executemany('''
            INSERT INTO productos (codigo, nombre, cantidad_por_caja)
            VALUES (?, ?, ?)
            ''', productos_ejemplo)

        # Crear la tabla ventas si no existe
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS ventas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT NOT NULL,
            nombre TEXT NOT NULL,
            cantidad INTEGER NOT NULL,
            fecha_carga TEXT NOT NULL
        )
        ''')

        # Crear índices para optimizar consultas
        self.cursor.execute('CREATE INDEX IF NOT EXISTS idx_fecha_carga ON ventas (fecha_carga)')
        self.cursor.execute('CREATE INDEX IF NOT EXISTS idx_codigo ON ventas (codigo)')
        self.conn.commit()

    def ejecutar_consulta(self, query, params=None):
        if params:
            self.cursor.execute(query, params)
        else:
            self.cursor.execute(query)
        return self.cursor.fetchall()

    def insertar_venta(self, codigo, nombre, cantidad, fecha_carga):
        self.cursor.execute('''
        INSERT INTO ventas (codigo, nombre, cantidad, fecha_carga)
        VALUES (?, ?, ?, ?)
        ''', (codigo, nombre, cantidad, fecha_carga))
        self.conn.commit()

    def cerrar(self):
        self.conn.close()