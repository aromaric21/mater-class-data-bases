CREATE DATABASE datalendocompany;

-- Création de la Tables clients
CREATE TABLE clients (
    client_id INT GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    nom TEXT NOT NULL,
    pays TEXT,
    date_creation DATE
);

-- Création de la Tables Produits
CREATE TABLE produits (
    product_id SERIAL PRIMARY KEY,
    nom TEXT NOT NULL,
    prix NUMERIC(10,2) NOT NULL
);

-- Création de la Tables Ventes
CREATE TABLE ventes (
    vente_id SERIAL PRIMARY KEY,
    client_id INTEGER REFERENCES clients(client_id),
    produit_id INTEGER REFERENCES produits(product_id),
    quantite INTEGER CHECK (quantite > 0),
    date_vente DATE NOT NULL
);

-- Insertion de quelques Clients
INSERT INTO clients (nom, pays, date_creation) 
VALUES
('Aminata', 'Sénégal', '2025-01-10'),
('Jean', 'France', '2025-02-01'),
('Fatou', 'Cameroun', '2025-01-15'),
('Luc', 'Belgique', '2025-02-05');

-- Insertion de quelques Produits
INSERT INTO produits (nom, prix) VALUES
('Ordinateur Portable', 1200),
('Casque Audio', 80),
('Souris USB', 30),
('Clé USB 32GB', 15);

-- Insertion de quelques Ventes
INSERT INTO ventes (client_id, produit_id, 
quantite, date_vente) VALUES
(1, 1, 1, '2025-02-10'),
(1, 2, 2, '2025-02-10'),
(2, 3, 1, '2025-02-11'),
(3, 2, 1, '2025-02-12'),
(4, 4, 3, '2025-02-12');


-- QUELQUES REQUÊTES SQL POUR ANALYSER LES DONNÉES 

-- 1. Lister tous les clients ?
SELECT * FROM clients;

-- 2. Lister tous les produits ?
SELECT * FROM produits;

-- 3. Lister tous les Ventes ?
SELECT * FROM ventes;

-- 5. Lister tous les produits premium(prix > 100)?
SELECT * FROM produits WHERE prix > 100;

-- 6. Combien avons-nous vendu au total ?
SELECT SUM(p.prix * v.quantite) AS total_ventes
FROM ventes v
JOIN produits p ON v.produit_id = p.product_id;

-- 6. Quels pays génèrent le plus de ventes ?
SELECT c.pays, COUNT(v.vente_id) AS nombre_ventes
FROM ventes v
JOIN clients c ON v.client_id = c.client_id
GROUP BY c.pays
ORDER BY nombre_ventes DESC;

-- 7. Quel est le pays qui génère le plus de ventes?
SELECT c.pays, COUNT(v.vente_id) AS nombre_ventes
FROM ventes v
JOIN clients c ON v.client_id = c.client_id
GROUP BY c.pays
ORDER BY nombre_ventes DESC
LIMIT 1;

--8. Le client Aminata a réalisé combien de ventes?
SELECT COUNT(v.vente_id) AS nombre_ventes_pour_aminata
FROM ventes v
JOIN clients c ON v.client_id = c.client_id
WHERE c.nom = 'Aminata';
