CREATE TABLE IF NOT EXISTS agentes (
    cf integer PRIMARY KEY AUTOINCREMENT NOT NULL,
    nombre varchar(20) NOT NULL,
    apellidos varchar(50) NOT NULL,
    categoria varchar(20) NOT NULL,
    residencia varchar(20)
);

CREATE TABLE IF NOT EXISTS turnos (
    turno varchar(10) PRIMARY KEY NOT NULL,
    inicio time,
    fin time,
    duracion time,
    categoria varchar(10) NOT NULL
);

CREATE TABLE IF NOT EXISTS turnos_agente (
    id integer PRIMARY KEY NOT NULL,
    turno varchar(10) NOT NULL,
    agente integer NOT NULL,
    fecha date NOT NULL,
    FOREIGN KEY(agente) REFERENCES agentes(cf),
    FOREIGN KEY(turno) REFERENCES turnos(turno)
);