import type { ProductoCatalogo } from './types';

export const MIBI_BRAND = {
  nombre: 'Mibi®',
  pieOficial: 'Mibi® · Ciudad de la Paz 2467, Belgrano, CABA · WhatsApp 11-5606-9384 · contacto@mibi.com.ar',
  web: 'https://www.mibi.com.ar/'
} as const;

export const VENDEDORES_MIBI = ['Juancho', 'Pablo Palacios', 'Agustín', 'Rocío', 'Administración'] as const;

export const CATALOGO_MIBI: ReadonlyArray<ProductoCatalogo> = [
  // 1. CHALECOS Y FLOTADORES
  {
    idGrupo: '6508-6509',
    nombre: 'Flotador Bracitos Monster (11 a 19 kg)',
    rubro: 'FLOTA',
    precioNeto: 23800,
    variantes: [
      { codigo: '6508', etiqueta: 'Green (Verde)', aliases: ['g', 'green', 'verde', '6508'] },
      { codigo: '6509', etiqueta: 'Pink (Rosa)', aliases: ['r', 'pink', 'rosa', '6509'] }
    ]
  },
  {
    idGrupo: '6500-6501',
    nombre: 'Chaleco Bubble Flotación Retirable (15 a 20 kg)',
    rubro: 'CH',
    precioNeto: 40661,
    variantes: [
      { codigo: '6500', etiqueta: 'Azul / Aqua', aliases: ['a', 'azul', 'aqua', 'celeste', '6500'] },
      { codigo: '6501', etiqueta: 'Pink (Rosa)', aliases: ['r', 'pink', 'rosa', '6501'] }
    ]
  },
  {
    idGrupo: '6502-6503',
    nombre: 'Chaleco Monster Placa Completa (3 a 6 años)',
    rubro: 'CH',
    precioNeto: 42645,
    variantes: [
      { codigo: '6502', etiqueta: 'Green (Verde)', aliases: ['g', 'green', 'verde', '6502'] },
      { codigo: '6503', etiqueta: 'Pink (Rosa)', aliases: ['r', 'pink', 'rosa', '6503'] }
    ]
  },
  {
    idGrupo: '6504-6505',
    nombre: 'Chaleco Shark Flotación Retirable',
    rubro: 'CH',
    precioNeto: 46612,
    variantes: [
      { codigo: '6504', etiqueta: 'T1 (1-3 años)', aliases: ['t1', '1 a 3', '1-3', 'talle 1', '6504'] },
      { codigo: '6505', etiqueta: 'T2 (3-6 años)', aliases: ['t2', '3 a 6', '3-6', 'talle 2', '6505'] }
    ]
  },
  {
    idGrupo: '6562-6563',
    nombre: 'Chaleco Coco Flotación Retirable',
    rubro: 'CH',
    precioNeto: 46612,
    variantes: [
      { codigo: '6562', etiqueta: 'T1 (1-3 años)', aliases: ['t1', '1 a 3', '1-3', 'talle 1', '6562'] },
      { codigo: '6563', etiqueta: 'T2 (3-6 años)', aliases: ['t2', '3 a 6', '3-6', 'talle 2', '6563'] }
    ]
  },
  {
    idGrupo: '6564-6565',
    nombre: 'Chaleco Rabb Flotación Retirable',
    rubro: 'CH',
    precioNeto: 46612,
    variantes: [
      { codigo: '6564', etiqueta: 'T1 (1-3 años)', aliases: ['t1', '1 a 3', '1-3', 'talle 1', '6564'] },
      { codigo: '6565', etiqueta: 'T2 (3-6 años)', aliases: ['t2', '3 a 6', '3-6', 'talle 2', '6565'] }
    ]
  },
  {
    idGrupo: '6566-6567',
    nombre: 'Chaleco Lovely Flotación Retirable',
    rubro: 'CH',
    precioNeto: 46612,
    variantes: [
      { codigo: '6566', etiqueta: 'T1 (1-3 años)', aliases: ['t1', '1 a 3', '1-3', 'talle 1', '6566'] },
      { codigo: '6567', etiqueta: 'T2 (3-6 años)', aliases: ['t2', '3 a 6', '3-6', 'talle 2', '6567'] }
    ]
  },
  {
    idGrupo: '6540-6568',
    nombre: 'Chaleco Bug Flotación Retirable',
    rubro: 'CH',
    precioNeto: 46612,
    variantes: [
      { codigo: '6540', etiqueta: 'T1 (1-3 años)', aliases: ['t1', '1 a 3', '1-3', 'talle 1', '6540'] },
      { codigo: '6568', etiqueta: 'T2 (3-6 años)', aliases: ['t2', '3 a 6', '3-6', 'talle 2', '6568'] }
    ]
  },
  {
    idGrupo: '6541-6569',
    nombre: 'Chaleco Margarita Flotación Retirable',
    rubro: 'CH',
    precioNeto: 46612,
    variantes: [
      { codigo: '6541', etiqueta: 'T1 (1-3 años)', aliases: ['t1', '1 a 3', '1-3', 'talle 1', '6541'] },
      { codigo: '6569', etiqueta: 'T2 (3-6 años)', aliases: ['t2', '3 a 6', '3-6', 'talle 2', '6569'] }
    ]
  },
  {
    idGrupo: '6539',
    nombre: 'Chaleco Petit Flotación Retirable',
    rubro: 'CH',
    precioNeto: 46612,
    variantes: [{ codigo: '6539', etiqueta: 'Único (1-3 años)', aliases: ['unico', '6539'] }]
  },
  {
    idGrupo: '6542',
    nombre: 'Chaleco Ocean Flotación Retirable',
    rubro: 'CH',
    precioNeto: 46612,
    variantes: [{ codigo: '6542', etiqueta: 'Único (1-3 años)', aliases: ['unico', '6542'] }]
  },
  {
    idGrupo: '6543',
    nombre: 'Chaleco Lila Splash Flotación Retirable',
    rubro: 'CH',
    precioNeto: 46612,
    variantes: [{ codigo: '6543', etiqueta: 'Único (1-3 años)', aliases: ['unico', '6543'] }]
  },
  {
    idGrupo: '6545',
    nombre: 'Chaleco Crabby Cool Flotación Retirable',
    rubro: 'CH',
    precioNeto: 46612,
    variantes: [{ codigo: '6545', etiqueta: 'Único (1-3 años)', aliases: ['unico', '6545'] }]
  },

  // 2. ANTIPARRAS Y MÁSCARAS
  {
    idGrupo: '6510',
    nombre: 'Antiparras Dino',
    rubro: 'ANTI',
    precioNeto: 9505,
    variantes: [{ codigo: '6510', etiqueta: 'Único Kids', aliases: ['unico', '6510'] }]
  },
  {
    idGrupo: '6511',
    nombre: 'Antiparras Pop',
    rubro: 'ANTI',
    precioNeto: 9505,
    variantes: [{ codigo: '6511', etiqueta: 'Único Kids', aliases: ['unico', '6511'] }]
  },
  {
    idGrupo: '6512',
    nombre: 'Antiparras Unicorn',
    rubro: 'ANTI',
    precioNeto: 10744,
    variantes: [{ codigo: '6512', etiqueta: 'Único Kids', aliases: ['unico', '6512'] }]
  },
  {
    idGrupo: '6513',
    nombre: 'Antiparras Sweet',
    rubro: 'ANTI',
    precioNeto: 10744,
    variantes: [{ codigo: '6513', etiqueta: 'Único Kids', aliases: ['unico', '6513'] }]
  },
  {
    idGrupo: '6528',
    nombre: 'Antiparras Candy',
    rubro: 'ANTI',
    precioNeto: 14050,
    variantes: [{ codigo: '6528', etiqueta: 'Único Kids', aliases: ['unico', '6528'] }]
  },
  {
    idGrupo: '6530',
    nombre: 'Antiparras Gummy',
    rubro: 'ANTI',
    precioNeto: 14050,
    variantes: [{ codigo: '6530', etiqueta: 'Único Kids', aliases: ['unico', '6530'] }]
  },
  {
    idGrupo: '6536',
    nombre: 'Antiparras Siren',
    rubro: 'ANTI',
    precioNeto: 17850,
    variantes: [{ codigo: '6536', etiqueta: 'Único Kids', aliases: ['unico', '6536'] }]
  },
  {
    idGrupo: '6537',
    nombre: 'Antiparras Shark',
    rubro: 'ANTI',
    precioNeto: 17850,
    variantes: [{ codigo: '6537', etiqueta: 'Único Kids', aliases: ['unico', '6537'] }]
  },
  {
    idGrupo: '6594',
    nombre: 'Máscara Snorkel Hipo',
    rubro: 'MASCA',
    precioNeto: 28100,
    variantes: [{ codigo: '6594', etiqueta: 'Único Kids', aliases: ['unico', '6594'] }]
  },
  {
    idGrupo: '6595',
    nombre: 'Máscara Snorkel Sharky',
    rubro: 'MASCA',
    precioNeto: 28100,
    variantes: [{ codigo: '6595', etiqueta: 'Único Kids', aliases: ['unico', '6595'] }]
  },

  // 3. ACCESORIOS DE AGUA (PAÑALES, ZAPATILLAS, MOCHILAS)
  {
    idGrupo: '6570',
    nombre: 'Pañal de Agua Shark (12-24m)',
    rubro: 'PAÑA',
    precioNeto: 18182,
    variantes: [{ codigo: '6570', etiqueta: 'Único (12-24m)', aliases: ['unico', '6570'] }]
  },
  {
    idGrupo: '6571',
    nombre: 'Pañal de Agua Rabb (12-24m)',
    rubro: 'PAÑA',
    precioNeto: 18182,
    variantes: [{ codigo: '6571', etiqueta: 'Único (12-24m)', aliases: ['unico', '6571'] }]
  },
  {
    idGrupo: '6572',
    nombre: 'Pañal de Agua Margarita (12-24m)',
    rubro: 'PAÑA',
    precioNeto: 18182,
    variantes: [{ codigo: '6572', etiqueta: 'Único (12-24m)', aliases: ['unico', '6572'] }]
  },
  {
    idGrupo: '6573',
    nombre: 'Pañal de Agua Coco (12-24m)',
    rubro: 'PAÑA',
    precioNeto: 18182,
    variantes: [{ codigo: '6573', etiqueta: 'Único (12-24m)', aliases: ['unico', '6573'] }]
  },
  {
    idGrupo: '6574',
    nombre: 'Pañal de Agua Bug (12-24m)',
    rubro: 'PAÑA',
    precioNeto: 18182,
    variantes: [{ codigo: '6574', etiqueta: 'Único (12-24m)', aliases: ['unico', '6574'] }]
  },
  {
    idGrupo: '6575',
    nombre: 'Pañal de Agua Lovely (12-24m)',
    rubro: 'PAÑA',
    precioNeto: 18182,
    variantes: [{ codigo: '6575', etiqueta: 'Único (12-24m)', aliases: ['unico', '6575'] }]
  },
  {
    idGrupo: '6576-6577',
    nombre: 'Zapatillas de Agua Shark',
    rubro: 'ZAPA',
    precioNeto: 16530,
    variantes: [
      { codigo: '6576', etiqueta: 'T1 (1-3 años)', aliases: ['t1', '1 a 3', '1-3', 'talle 1', '6576'] },
      { codigo: '6577', etiqueta: 'T2 (3-6 años)', aliases: ['t2', '3 a 6', '3-6', 'talle 2', '6577'] }
    ]
  },
  {
    idGrupo: '6578-6579',
    nombre: 'Zapatillas de Agua Rabb',
    rubro: 'ZAPA',
    precioNeto: 16530,
    variantes: [
      { codigo: '6578', etiqueta: 'T1 (1-3 años)', aliases: ['t1', '1 a 3', '1-3', 'talle 1', '6578'] },
      { codigo: '6579', etiqueta: 'T2 (3-6 años)', aliases: ['t2', '3 a 6', '3-6', 'talle 2', '6579'] }
    ]
  },
  {
    idGrupo: '6580-6581',
    nombre: 'Zapatillas de Agua Margarita',
    rubro: 'ZAPA',
    precioNeto: 16530,
    variantes: [
      { codigo: '6580', etiqueta: 'T1 (1-3 años)', aliases: ['t1', '1 a 3', '1-3', 'talle 1', '6580'] },
      { codigo: '6581', etiqueta: 'T2 (3-6 años)', aliases: ['t2', '3 a 6', '3-6', 'talle 2', '6581'] }
    ]
  },
  {
    idGrupo: '6582-6583',
    nombre: 'Zapatillas de Agua Coco',
    rubro: 'ZAPA',
    precioNeto: 16530,
    variantes: [
      { codigo: '6582', etiqueta: 'T1 (1-3 años)', aliases: ['t1', '1 a 3', '1-3', 'talle 1', '6582'] },
      { codigo: '6583', etiqueta: 'T2 (3-6 años)', aliases: ['t2', '3 a 6', '3-6', 'talle 2', '6583'] }
    ]
  },
  {
    idGrupo: '6584-6585',
    nombre: 'Zapatillas de Agua Bug',
    rubro: 'ZAPA',
    precioNeto: 16530,
    variantes: [
      { codigo: '6584', etiqueta: 'T1 (1-3 años)', aliases: ['t1', '1 a 3', '1-3', 'talle 1', '6584'] },
      { codigo: '6585', etiqueta: 'T2 (3-6 años)', aliases: ['t2', '3 a 6', '3-6', 'talle 2', '6585'] }
    ]
  },
  {
    idGrupo: '6586-6587',
    nombre: 'Zapatillas de Agua Lovely',
    rubro: 'ZAPA',
    precioNeto: 16530,
    variantes: [
      { codigo: '6586', etiqueta: 'T1 (1-3 años)', aliases: ['t1', '1 a 3', '1-3', 'talle 1', '6586'] },
      { codigo: '6587', etiqueta: 'T2 (3-6 años)', aliases: ['t2', '3 a 6', '3-6', 'talle 2', '6587'] }
    ]
  },
  {
    idGrupo: '6588',
    nombre: 'Mochila de Arena Shark',
    rubro: 'MOCHI',
    precioNeto: 20662,
    variantes: [{ codigo: '6588', etiqueta: 'Único Kids', aliases: ['unico', '6588'] }]
  },
  {
    idGrupo: '6589',
    nombre: 'Mochila de Arena Rabb',
    rubro: 'MOCHI',
    precioNeto: 20662,
    variantes: [{ codigo: '6589', etiqueta: 'Único Kids', aliases: ['unico', '6589'] }]
  },
  {
    idGrupo: '6590',
    nombre: 'Mochila de Arena Coco',
    rubro: 'MOCHI',
    precioNeto: 20662,
    variantes: [{ codigo: '6590', etiqueta: 'Único Kids', aliases: ['unico', '6590'] }]
  },
  {
    idGrupo: '6591',
    nombre: 'Mochila de Arena Margarita',
    rubro: 'MOCHI',
    precioNeto: 20662,
    variantes: [{ codigo: '6591', etiqueta: 'Único Kids', aliases: ['unico', '6591'] }]
  },
  {
    idGrupo: '6592',
    nombre: 'Mochila de Arena Bug',
    rubro: 'MOCHI',
    precioNeto: 20662,
    variantes: [{ codigo: '6592', etiqueta: 'Único Kids', aliases: ['unico', '6592'] }]
  },
  {
    idGrupo: '6593',
    nombre: 'Mochila de Arena Lovely',
    rubro: 'MOCHI',
    precioNeto: 20662,
    variantes: [{ codigo: '6593', etiqueta: 'Único Kids', aliases: ['unico', '6593'] }]
  },

  // 4. INFLABLES Y FLOTADORES
  {
    idGrupo: '6602',
    nombre: 'Inflable Sunny Pink',
    rubro: 'INFLA',
    precioNeto: 30580,
    variantes: [{ codigo: '6602', etiqueta: 'Único', aliases: ['unico', '6602'] }]
  },
  {
    idGrupo: '6603',
    nombre: 'Inflable Sunny Aqua',
    rubro: 'INFLA',
    precioNeto: 30580,
    variantes: [{ codigo: '6603', etiqueta: 'Único', aliases: ['unico', '6603'] }]
  },
  {
    idGrupo: '6604',
    nombre: 'Inflable Wave Pink',
    rubro: 'INFLA',
    precioNeto: 27275,
    variantes: [{ codigo: '6604', etiqueta: 'Único', aliases: ['unico', '6604'] }]
  },
  {
    idGrupo: '6605',
    nombre: 'Inflable Wave Blue',
    rubro: 'INFLA',
    precioNeto: 27275,
    variantes: [{ codigo: '6605', etiqueta: 'Único', aliases: ['unico', '6605'] }]
  },
  {
    idGrupo: '6606',
    nombre: 'Flotador Bunny Hug',
    rubro: 'FLOTA',
    precioNeto: 37190,
    variantes: [{ codigo: '6606', etiqueta: 'Único', aliases: ['unico', '6606'] }]
  },
  {
    idGrupo: '6607',
    nombre: 'Flotador Dino Hug',
    rubro: 'FLOTA',
    precioNeto: 37190,
    variantes: [{ codigo: '6607', etiqueta: 'Único', aliases: ['unico', '6607'] }]
  },
  {
    idGrupo: '6608',
    nombre: 'Pileta Love Pool',
    rubro: 'INFLA',
    precioNeto: 28925,
    variantes: [{ codigo: '6608', etiqueta: 'Único', aliases: ['unico', '6608'] }]
  },
  {
    idGrupo: '6609',
    nombre: 'Pileta Baby Dino',
    rubro: 'INFLA',
    precioNeto: 29752,
    variantes: [{ codigo: '6609', etiqueta: 'Único', aliases: ['unico', '6609'] }]
  },

  // 5. VARIOS Y PACKAGING
  {
    idGrupo: '6610',
    nombre: 'Pelela Lola V1',
    rubro: 'PELELAS',
    precioNeto: 33060,
    variantes: [{ codigo: '6610', etiqueta: 'Único', aliases: ['unico', '6610'] }]
  },
  {
    idGrupo: '6611',
    nombre: 'Pelela Lola V2',
    rubro: 'PELELAS',
    precioNeto: 33060,
    variantes: [{ codigo: '6611', etiqueta: 'Único', aliases: ['unico', '6611'] }]
  },
  {
    idGrupo: '6612',
    nombre: 'Bolsa Troquelada 33x35',
    rubro: 'PACK',
    precioNeto: 3000,
    variantes: [{ codigo: '6612', etiqueta: 'Único', aliases: ['unico', '6612'] }]
  },
  {
    idGrupo: '6613',
    nombre: 'Bolso Playero con Fuelle',
    rubro: 'PACK',
    precioNeto: 7000,
    variantes: [{ codigo: '6613', etiqueta: 'Único', aliases: ['unico', '6613'] }]
  },
  {
    idGrupo: '6614',
    nombre: 'Bolsa PVC 28x28',
    rubro: 'PACK',
    precioNeto: 3500,
    variantes: [{ codigo: '6614', etiqueta: 'Único', aliases: ['unico', '6614'] }]
  }
];

export const CATALOGO_BY_GRUPO = new Map<string, ProductoCatalogo>();
export const CATALOGO_BY_CODIGO = new Map<string, ProductoCatalogo>();

for (const prod of CATALOGO_MIBI) {
  CATALOGO_BY_GRUPO.set(prod.idGrupo, prod);
  for (const v of prod.variantes) {
    CATALOGO_BY_CODIGO.set(v.codigo, prod);
  }
}

// Artículos complementarios presentes en las filas 74–108 de LISTA N° 006 (Ponchos, Toallas, Pilusos y Gorras)
export const CATALOGO_COMPLEMENTARIO_006: ReadonlyArray<ProductoCatalogo> = [
  // PONCHOS
  {
    idGrupo: '6514',
    nombre: 'Poncho Shark (60x60)',
    rubro: 'PONCHOS',
    precioNeto: 21487.6,
    variantes: [{ codigo: '6514', etiqueta: 'Talle Único', aliases: ['unico', '6514'] }]
  },
  {
    idGrupo: '6515',
    nombre: 'Poncho Octopus (60x60)',
    rubro: 'PONCHOS',
    precioNeto: 21487.6,
    variantes: [{ codigo: '6515', etiqueta: 'Talle Único', aliases: ['unico', '6515'] }]
  },
  {
    idGrupo: '6516-6517',
    nombre: 'Poncho Monsters (60x60)',
    rubro: 'PONCHOS',
    precioNeto: 21487.6,
    variantes: [
      { codigo: '6516', etiqueta: 'Green (Verde)', aliases: ['g', 'green', 'verde', '6516'] },
      { codigo: '6517', etiqueta: 'Pink (Rosa)', aliases: ['p', 'r', 'pink', 'rosa', '6517'] }
    ]
  },
  {
    idGrupo: '6518-6519',
    nombre: 'Poncho Bubble (60x60)',
    rubro: 'PONCHOS',
    precioNeto: 21487.6,
    variantes: [
      { codigo: '6518', etiqueta: 'Green (Verde)', aliases: ['g', 'green', 'verde', '6518'] },
      { codigo: '6519', etiqueta: 'Pink (Rosa)', aliases: ['p', 'r', 'pink', 'rosa', '6519'] }
    ]
  },
  {
    idGrupo: '6546',
    nombre: 'Poncho Monster Neon (60x60)',
    rubro: 'PONCHOS',
    precioNeto: 23966.94,
    variantes: [{ codigo: '6546', etiqueta: 'Talle Único', aliases: ['unico', '6546'] }]
  },
  {
    idGrupo: '6551',
    nombre: 'Poncho Lila Splash (60x60)',
    rubro: 'PONCHOS',
    precioNeto: 23966.94,
    variantes: [{ codigo: '6551', etiqueta: 'Talle Único', aliases: ['unico', '6551'] }]
  },
  {
    idGrupo: '6599',
    nombre: 'Poncho Rabb (60x60)',
    rubro: 'PONCHOS',
    precioNeto: 23966.94,
    variantes: [{ codigo: '6599', etiqueta: 'Talle Único', aliases: ['unico', '6599'] }]
  },
  {
    idGrupo: '6600',
    nombre: 'Poncho Coco (60x60)',
    rubro: 'PONCHOS',
    precioNeto: 23966.94,
    variantes: [{ codigo: '6600', etiqueta: 'Talle Único', aliases: ['unico', '6600'] }]
  },
  {
    idGrupo: '6601',
    nombre: 'Poncho Lovely (60x60)',
    rubro: 'PONCHOS',
    precioNeto: 23966.94,
    variantes: [{ codigo: '6601', etiqueta: 'Talle Único', aliases: ['unico', '6601'] }]
  },
  {
    idGrupo: '6552',
    nombre: 'Poncho Crab Pop (60x60)',
    rubro: 'PONCHOS',
    precioNeto: 25785,
    variantes: [{ codigo: '6552', etiqueta: 'Talle Único', aliases: ['unico', '6552'] }]
  },

  // TOALLAS MICROFIBRA
  {
    idGrupo: '6520-6521',
    nombre: 'Toallas Bubble Microfibra (1,50x0,90)',
    rubro: 'TOALLAS',
    precioNeto: 18595.04,
    variantes: [
      { codigo: '6520', etiqueta: 'Blue (Azul / P)', aliases: ['b', 'blue', 'azul', 'p', '6520'] },
      { codigo: '6521', etiqueta: 'Pink (Rosa / B)', aliases: ['pink', 'rosa', 'r', 'b', '6521'] }
    ]
  },
  {
    idGrupo: '6522',
    nombre: 'Toalla Shark Microfibra (1,50x0,90)',
    rubro: 'TOALLAS',
    precioNeto: 18595.04,
    variantes: [{ codigo: '6522', etiqueta: 'Talle Único', aliases: ['unico', '6522'] }]
  },
  {
    idGrupo: '6523',
    nombre: 'Toalla Octopus Microfibra (1,50x0,90)',
    rubro: 'TOALLAS',
    precioNeto: 18181.81,
    variantes: [{ codigo: '6523', etiqueta: 'Talle Único', aliases: ['unico', '6523'] }]
  },
  {
    idGrupo: '6524-6525',
    nombre: 'Toallas Monsters Microfibra (1,50x0,90)',
    rubro: 'TOALLAS',
    precioNeto: 18595.04,
    variantes: [
      { codigo: '6524', etiqueta: 'Green / Pink', aliases: ['g', 'green', 'verde', 'p', '6524'] },
      { codigo: '6525', etiqueta: 'Blue / Pink', aliases: ['b', 'blue', 'azul', 'pink', 'rosa', '6525'] }
    ]
  },
  {
    idGrupo: '6554',
    nombre: 'Toalla Monster Neon Microfibra (1,50x0,70)',
    rubro: 'TOALLAS',
    precioNeto: 18595.04,
    variantes: [{ codigo: '6554', etiqueta: 'Talle Único', aliases: ['unico', '6554'] }]
  },
  {
    idGrupo: '6555',
    nombre: 'Toalla Petit Microfibra (1,50x0,90)',
    rubro: 'TOALLAS',
    precioNeto: 18595.04,
    variantes: [{ codigo: '6555', etiqueta: 'Talle Único', aliases: ['unico', '6555'] }]
  },
  {
    idGrupo: '6556',
    nombre: 'Toalla Bug Microfibra (1,50x0,90)',
    rubro: 'TOALLAS',
    precioNeto: 18595.04,
    variantes: [{ codigo: '6556', etiqueta: 'Talle Único', aliases: ['unico', '6556'] }]
  },
  {
    idGrupo: '6557',
    nombre: 'Toalla Margarita Microfibra (1,50x0,90)',
    rubro: 'TOALLAS',
    precioNeto: 18595.04,
    variantes: [{ codigo: '6557', etiqueta: 'Talle Único', aliases: ['unico', '6557'] }]
  },
  {
    idGrupo: '6558',
    nombre: 'Toalla Ocean Microfibra (1,50x0,90)',
    rubro: 'TOALLAS',
    precioNeto: 18595.04,
    variantes: [{ codigo: '6558', etiqueta: 'Talle Único', aliases: ['unico', '6558'] }]
  },
  {
    idGrupo: '6559',
    nombre: 'Toalla Lila Splash Microfibra (1,50x0,90)',
    rubro: 'TOALLAS',
    precioNeto: 18595.04,
    variantes: [{ codigo: '6559', etiqueta: 'Talle Único', aliases: ['unico', '6559'] }]
  },
  {
    idGrupo: '6560',
    nombre: 'Toalla Crab Pop Microfibra (1,50x0,90)',
    rubro: 'TOALLAS',
    precioNeto: 18595.04,
    variantes: [{ codigo: '6560', etiqueta: 'Talle Único', aliases: ['unico', '6560'] }]
  },
  {
    idGrupo: '6561',
    nombre: 'Toalla Crabby Cool Microfibra (1,50x0,90)',
    rubro: 'TOALLAS',
    precioNeto: 18595.04,
    variantes: [{ codigo: '6561', etiqueta: 'Talle Único', aliases: ['unico', '6561'] }]
  },
  {
    idGrupo: '6596',
    nombre: 'Toalla Rabb Microfibra (1,50x0,90)',
    rubro: 'TOALLAS',
    precioNeto: 18595.04,
    variantes: [{ codigo: '6596', etiqueta: 'Talle Único', aliases: ['unico', '6596'] }]
  },
  {
    idGrupo: '6597',
    nombre: 'Toalla Coco Microfibra (1,50x0,90)',
    rubro: 'TOALLAS',
    precioNeto: 18595.04,
    variantes: [{ codigo: '6597', etiqueta: 'Talle Único', aliases: ['unico', '6597'] }]
  },
  {
    idGrupo: '6598',
    nombre: 'Toalla Lovely Microfibra (1,50x0,90)',
    rubro: 'TOALLAS',
    precioNeto: 18595.04,
    variantes: [{ codigo: '6598', etiqueta: 'Talle Único', aliases: ['unico', '6598'] }]
  },

  // PILUSOS Y GORRAS (TALLE ÚNICO)
  {
    idGrupo: '6526',
    nombre: 'Pilusos Gabardina',
    rubro: 'PILUSOS',
    precioNeto: 6942.14,
    variantes: [{ codigo: '6526', etiqueta: 'Talle Único', aliases: ['unico', 'surtido', 'varios', '6526'] }]
  },
  {
    idGrupo: '6527',
    nombre: 'Gorras Polyester con Red',
    rubro: 'GORRAS',
    precioNeto: 6942.14,
    variantes: [{ codigo: '6527', etiqueta: 'Talle Único', aliases: ['unico', 'surtido', '6527'] }]
  }
];

export const CATALOGO_EXTENDIDO_MIBI: ReadonlyArray<ProductoCatalogo> = [
  ...CATALOGO_MIBI,
  ...CATALOGO_COMPLEMENTARIO_006
];

export const CATALOGO_EXTENDIDO_BY_GRUPO = new Map<string, ProductoCatalogo>();
export const CATALOGO_EXTENDIDO_BY_CODIGO = new Map<string, ProductoCatalogo>();

for (const prod of CATALOGO_EXTENDIDO_MIBI) {
  CATALOGO_EXTENDIDO_BY_GRUPO.set(prod.idGrupo, prod);
  for (const v of prod.variantes) {
    CATALOGO_EXTENDIDO_BY_CODIGO.set(v.codigo, prod);
  }
}
