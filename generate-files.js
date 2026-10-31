const XLSX = require('xlsx');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, AlignmentType } = require('docx');
const fs = require('fs');

// ==================== KEYWORD DATA ====================

const primaryKeywords = [
  { keyword: 'wholesale products from China to Bangladesh', intent: 'Commercial/Transactional', priority: 'Critical', difficulty: 'Medium', targetUrl: '/' },
  { keyword: 'China to Bangladesh wholesale import', intent: 'Commercial', priority: 'Critical', difficulty: 'Medium', targetUrl: '/' },
  { keyword: 'wholesale sourcing China Bangladesh', intent: 'Commercial', priority: 'Critical', difficulty: 'Medium', targetUrl: '/products' },
  { keyword: 'buy wholesale from China to Bangladesh', intent: 'Transactional', priority: 'Critical', difficulty: 'Medium', targetUrl: '/products' },
  { keyword: 'China wholesale Bangladesh supplier', intent: 'Commercial', priority: 'High', difficulty: 'Medium', targetUrl: '/products?category=' },
  { keyword: 'import products from China to Bangladesh', intent: 'Informational/Commercial', priority: 'High', difficulty: 'Medium', targetUrl: '/about' },
  { keyword: 'wholesale marketplace Bangladesh', intent: 'Commercial', priority: 'High', difficulty: 'Medium', targetUrl: '/products' },
  { keyword: 'China sourcing agent Bangladesh', intent: 'Commercial', priority: 'High', difficulty: 'Medium', targetUrl: '/about' },
  { keyword: 'Bangladesh China import business', intent: 'Commercial', priority: 'High', difficulty: 'Medium', targetUrl: '/about' },
  { keyword: 'wholesale products Bangladesh', intent: 'Commercial', priority: 'High', difficulty: 'High', targetUrl: '/products' },
];

const secondaryKeywords = [
  { keyword: '1688 to Bangladesh sourcing', intent: 'Commercial', volume: 'High', targetUrl: '/products' },
  { keyword: 'Alibaba to Bangladesh shipping', intent: 'Commercial', volume: 'High', targetUrl: '/shipping' },
  { keyword: 'China factory direct Bangladesh', intent: 'Commercial', volume: 'Medium', targetUrl: '/products' },
  { keyword: 'China supplier verification Bangladesh', intent: 'Informational', volume: 'Medium', targetUrl: '/about' },
  { keyword: 'product sourcing from China Bangladesh', intent: 'Commercial', volume: 'High', targetUrl: '/products' },
  { keyword: 'China wholesale agent Dhaka', intent: 'Commercial', volume: 'Medium', targetUrl: '/contact' },
  { keyword: 'buy from 1688 Bangladesh', intent: 'Transactional', volume: 'Medium', targetUrl: '/products' },
  { keyword: 'Alibaba Bangladesh shipping cost', intent: 'Commercial', volume: 'Medium', targetUrl: '/calculator' },
  { keyword: 'China to Bangladesh freight forwarder', intent: 'Commercial', volume: 'Medium', targetUrl: '/shipping' },
  { keyword: 'import license Bangladesh China', intent: 'Informational', volume: 'Low', targetUrl: '/faq' },
  { keyword: 'air freight China to Bangladesh', intent: 'Commercial', volume: 'High', targetUrl: '/shipping' },
  { keyword: 'sea freight China to Bangladesh', intent: 'Commercial', volume: 'High', targetUrl: '/shipping' },
  { keyword: 'shipping rate China Bangladesh', intent: 'Commercial', volume: 'Critical', targetUrl: '/calculator' },
  { keyword: 'door to door delivery China Bangladesh', intent: 'Commercial', volume: 'High', targetUrl: '/shipping' },
  { keyword: 'customs clearance Bangladesh import', intent: 'Informational', volume: 'Medium', targetUrl: '/faq' },
  { keyword: 'Bangladesh import duty calculator', intent: 'Commercial', volume: 'Medium', targetUrl: '/calculator' },
  { keyword: 'freight forwarder Dhaka China', intent: 'Commercial', volume: 'Medium', targetUrl: '/contact' },
  { keyword: 'parcel shipping China Bangladesh', intent: 'Transactional', volume: 'Medium', targetUrl: '/shipping' },
];

const categoryKeywords = [
  { category: 'Shoes', id: '1038378', primary: 'wholesale shoes Bangladesh', lsi: 'China shoes import Dhaka, footwear wholesale Bangladesh', url: '/products?category=1038378' },
  { category: 'Toys', id: '1813', primary: 'wholesale toys Bangladesh', lsi: 'China toys import, kids toys wholesale Dhaka', url: '/products?category=1813' },
  { category: 'Beauty & Cosmetics', id: '97', primary: 'wholesale cosmetics Bangladesh', lsi: 'beauty products import China, skincare wholesale Dhaka', url: '/products?category=97' },
  { category: 'Electronics & Computer', id: '7', primary: 'wholesale electronics Bangladesh', lsi: 'China electronics import Dhaka, mobile accessories wholesale', url: '/products?category=7' },
  { category: "Women's Clothing", id: '10166', primary: 'wholesale clothing Bangladesh', lsi: "women's clothing China import, fashion wholesale Dhaka", url: '/products?category=10166' },
  { category: 'Home & Kitchen', id: '15', primary: 'wholesale home kitchen Bangladesh', lsi: 'kitchenware import China, home essentials wholesale Dhaka', url: '/products?category=15' },
  { category: 'Sports & Outdoor', id: '18', primary: 'wholesale sports equipment Bangladesh', lsi: 'fitness gear import China, outdoor gear wholesale Dhaka', url: '/products?category=18' },
  { category: 'Luggage & Leather Goods', id: '1042954', primary: 'wholesale bags Bangladesh', lsi: 'leather goods China import, luggage wholesale Dhaka', url: '/products?category=1042954' },
  { category: 'Agriculture', id: '1', primary: 'wholesale agriculture Bangladesh', lsi: 'farming equipment China import, agri supplies wholesale Dhaka', url: '/products?category=1' },
  { category: 'Lighting', id: '58', primary: 'wholesale lighting Bangladesh', lsi: 'LED lights China import, lighting fixtures wholesale Dhaka', url: '/products?category=58' },
  { category: 'Household Appliances', id: '6', primary: 'wholesale appliances Bangladesh', lsi: 'home appliances China import, electrical goods wholesale Dhaka', url: '/products?category=6' },
  { category: 'Pets & Gardening', id: '122916001', primary: 'wholesale pet supplies Bangladesh', lsi: 'pet products China import, gardening supplies wholesale Dhaka', url: '/products?category=122916001' },
  { category: 'Digital & Computer', id: '7', primary: 'wholesale computer accessories Bangladesh', lsi: 'IT accessories China import, digital products wholesale Dhaka', url: '/products?category=7' },
  { category: "Men's Clothing", id: '10165', primary: 'wholesale menswear Bangladesh', lsi: "men's clothing China import, fashion wholesale Dhaka", url: '/products?category=10165' },
  { category: "Children's Clothing", id: '311', primary: 'wholesale kids clothing Bangladesh', lsi: "children's wear China import, baby clothes wholesale Dhaka", url: '/products?category=311' },
];

const bengaliKeywords = [
  { bengali: 'চীন থেকে পণ্য আমদানি', english: 'Import products from China', intent: 'Commercial', targetUrl: '/about' },
  { bengali: 'চীন থেকে পাইকারি পণ্য', english: 'Wholesale products from China', intent: 'Commercial', targetUrl: '/products' },
  { bengali: 'চীনা পণ্য সোর্সিং বাংলাদেশ', english: 'China product sourcing Bangladesh', intent: 'Commercial', targetUrl: '/products' },
  { bengali: 'চীন থেকে শিপিং খরচ', english: 'Shipping cost from China', intent: 'Commercial', targetUrl: '/calculator' },
  { bengali: 'বাংলাদেশে পাইকারি বাজার', english: 'Wholesale market Bangladesh', intent: 'Commercial', targetUrl: '/products' },
  { bengali: 'চীনা সরবরাহকারী বাংলাদেশ', english: 'China supplier Bangladesh', intent: 'Commercial', targetUrl: '/about' },
  { bengali: 'চীন থেকে আমদানি গাইড', english: 'Import from China guide', intent: 'Informational', targetUrl: '/faq' },
];

const longTailKeywords = [
  { keyword: 'how to import small quantity from China to Bangladesh', intent: 'Informational', difficulty: 'Low', targetUrl: '/blog' },
  { keyword: 'best products to import from China to Bangladesh 2026', intent: 'Informational', difficulty: 'Low', targetUrl: '/blog' },
  { keyword: '1688.com buying guide Bangladesh buyers', intent: 'Informational', difficulty: 'Low', targetUrl: '/blog' },
  { keyword: 'China to Bangladesh air freight per kg rate 2026', intent: 'Commercial', difficulty: 'Medium', targetUrl: '/shipping' },
  { keyword: 'sea freight LCL China to Chittagong cost', intent: 'Commercial', difficulty: 'Medium', targetUrl: '/shipping' },
  { keyword: 'Bangladesh import duty rate for electronics from China', intent: 'Informational', difficulty: 'Low', targetUrl: '/faq' },
  { keyword: 'wholesale mobile accessories Bangladesh import', intent: 'Commercial', difficulty: 'Medium', targetUrl: '/products?category=7' },
  { keyword: "women's fashion wholesale China to Dhaka", intent: 'Commercial', difficulty: 'Medium', targetUrl: '/products?category=10166' },
  { keyword: 'cosmetic import Bangladesh BSTI clearance', intent: 'Informational', difficulty: 'Low', targetUrl: '/faq' },
  { keyword: 'China factory wholesale Bangladesh no MOQ', intent: 'Commercial', difficulty: 'Low', targetUrl: '/products' },
  { keyword: 'dropshipping China to Bangladesh', intent: 'Commercial', difficulty: 'Medium', targetUrl: '/about' },
  { keyword: 'Amazon FBA prep China Bangladesh', intent: 'Commercial', difficulty: 'Low', targetUrl: '/blog' },
  { keyword: 'Shopify sourcing China Bangladesh', intent: 'Commercial', difficulty: 'Low', targetUrl: '/blog' },
  { keyword: 'Facebook live selling products China import Bangladesh', intent: 'Commercial', difficulty: 'Low', targetUrl: '/blog' },
  { keyword: 'Daraz seller sourcing China Bangladesh', intent: 'Commercial', difficulty: 'Low', targetUrl: '/blog' },
];

// ==================== XLSX GENERATION ====================

function createXLSX() {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Primary Keywords
  const ws1 = XLSX.utils.json_to_sheet(primaryKeywords.map(k => ({
    Keyword: k.keyword,
    Intent: k.intent,
    Priority: k.priority,
    Difficulty: k.difficulty,
    Target_URL: k.targetUrl
  })));
  XLSX.utils.book_append_sheet(wb, ws1, '1_Primary_Keywords');

  // Sheet 2: Secondary Keywords
  const ws2 = XLSX.utils.json_to_sheet(secondaryKeywords.map(k => ({
    Keyword: k.keyword,
    Intent: k.intent,
    Volume_Tier: k.volume,
    Target_URL: k.targetUrl
  })));
  XLSX.utils.book_append_sheet(wb, ws2, '2_Secondary_Keywords');

  // Sheet 3: Category Keywords
  const ws3 = XLSX.utils.json_to_sheet(categoryKeywords.map(k => ({
    Category: k.category,
    Category_ID: k.id,
    Primary_Keyword: k.primary,
    LSI_Keywords: k.lsi,
    Target_URL: k.url
  })));
  XLSX.utils.book_append_sheet(wb, ws3, '3_Category_Keywords');

  // Sheet 4: Bengali Keywords
  const ws4 = XLSX.utils.json_to_sheet(bengaliKeywords.map(k => ({
    Bengali_Keyword: k.bengali,
    English_Translation: k.english,
    Intent: k.intent,
    Target_URL: k.targetUrl
  })));
  XLSX.utils.book_append_sheet(wb, ws4, '4_Bengali_Keywords');

  // Sheet 5: Long Tail Keywords
  const ws5 = XLSX.utils.json_to_sheet(longTailKeywords.map(k => ({
    Keyword: k.keyword,
    Intent: k.intent,
    Difficulty: k.difficulty,
    Target_URL: k.targetUrl
  })));
  XLSX.utils.book_append_sheet(wb, ws5, '5_Long_Tail_Keywords');

  XLSX.writeFile(wb, '/workspace/app/WholesalePlus_Keyword_Research.xlsx');
  console.log('XLSX created successfully');
}

// ==================== DOCX GENERATION ====================

function createDOCX() {
  const doc = new Document({
    title: 'Wholesale Plus - Keyword Research, Clustering & Mapping',
    author: 'Kilo AI SEO Audit',
    sections: [{
      properties: {},
      children: [
        // Title
        new Paragraph({
          text: 'Wholesale Plus - Comprehensive Keyword Research, Clustering & Mapping',
          heading: HeadingLevel.TITLE,
          alignment: AlignmentType.CENTER,
        }),
        new Paragraph({
          text: 'Target Website: https://wholesaleplus.com.bd/\nMarket: Bangladesh (B2B Cross-Border E-commerce)\nResearch Date: October 31, 2026',
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
        }),

        new Paragraph({
          text: '1. PRIMARY KEYWORDS (High Volume, Commercial Intent)',
          heading: HeadingLevel.HEADING_1,
          spacing: { before: 200, after: 100 },
        }),
        ...createTable(
          ['#', 'Keyword', 'Intent', 'Priority', 'Difficulty', 'Target URL'],
          primaryKeywords.map((k, i) => [i + 1, k.keyword, k.intent, k.priority, k.difficulty, k.targetUrl])
        ),

        new Paragraph({
          text: '2. SECONDARY & LSI KEYWORDS',
          heading: HeadingLevel.HEADING_1,
          spacing: { before: 200, after: 100 },
        }),
        new Paragraph({
          text: '2.1 Sourcing & Import Process Keywords',
          heading: HeadingLevel.HEADING_2,
          spacing: { before: 100, after: 50 },
        }),
        ...createTable(
          ['Keyword', 'Intent', 'Volume Tier', 'Target Page'],
          secondaryKeywords.filter(k => !k.keyword.includes('freight') && !k.keyword.includes('shipping') && !k.keyword.includes('door')).map(k => [k.keyword, k.intent, k.volume, k.targetUrl])
        ),

        new Paragraph({
          text: '2.2 Shipping & Logistics Keywords',
          heading: HeadingLevel.HEADING_2,
          spacing: { before: 100, after: 50 },
        }),
        ...createTable(
          ['Keyword', 'Intent', 'Volume Tier', 'Target Page'],
          secondaryKeywords.filter(k => k.keyword.includes('freight') || k.keyword.includes('shipping') || k.keyword.includes('door') || k.keyword.includes('customs') || k.keyword.includes('duty') || k.keyword.includes('parcel') || k.keyword.includes('forwarder')).map(k => [k.keyword, k.intent, k.volume, k.targetUrl])
        ),

        new Paragraph({
          text: '3. PRODUCT CATEGORY KEYWORDS',
          heading: HeadingLevel.HEADING_1,
          spacing: { before: 200, after: 100 },
        }),
        ...createTable(
          ['Category', 'Category ID', 'Primary Keyword', 'LSI Keywords', 'Target URL'],
          categoryKeywords.map(k => [k.category, k.id, k.primary, k.lsi, k.url])
        ),

        new Paragraph({
          text: '4. BENGALI LANGUAGE KEYWORDS (Local Market)',
          heading: HeadingLevel.HEADING_1,
          spacing: { before: 200, after: 100 },
        }),
        ...createTable(
          ['Bengali Keyword', 'English Translation', 'Intent', 'Target Page'],
          bengaliKeywords.map(k => [k.bengali, k.english, k.intent, k.targetUrl])
        ),

        new Paragraph({
          text: '5. LONG-TAIL KEYWORD OPPORTUNITIES (Low Competition, High Intent)',
          heading: HeadingLevel.HEADING_1,
          spacing: { before: 200, after: 100 },
        }),
        ...createTable(
          ['Long-Tail Keyword', 'Intent', 'Difficulty', 'Target Page'],
          longTailKeywords.map(k => [k.keyword, k.intent, k.difficulty, k.targetUrl])
        ),

        new Paragraph({
          text: '6. KEYWORD CLUSTERS & CONTENT STRATEGY',
          heading: HeadingLevel.HEADING_1,
          spacing: { before: 200, after: 100 },
        }),
        new Paragraph({
          text: 'Cluster 1: Core Sourcing (Highest Priority)\nTheme: Finding and buying wholesale products from China for Bangladesh market\nTarget Pages: Homepage + Products page optimization, category landing pages',
          spacing: { after: 100 },
        }),
        new Paragraph({
          text: 'Cluster 2: Import Process & Logistics\nTheme: How to import from China, shipping, customs, costs\nTarget Pages: Shipping page, Calculator page, FAQ page, Blog content',
          spacing: { after: 100 },
        }),
        new Paragraph({
          text: 'Cluster 3: Product Categories\nTheme: Specific product categories for wholesale import\nTarget Pages: Category landing pages with unique content for each',
          spacing: { after: 100 },
        }),
        new Paragraph({
          text: 'Cluster 4: Sourcing Agents & Platforms\nTheme: China sourcing agents, 1688, Alibaba for Bangladesh\nTarget Pages: About page, Blog comparison articles',
          spacing: { after: 100 },
        }),
        new Paragraph({
          text: 'Cluster 5: Local Bengali Keywords\nTheme: Bengali-language search terms for local audience\nTarget Pages: Bengali versions of key pages, hreflang tags',
          spacing: { after: 200 },
        }),

        new Paragraph({
          text: '7. IMPLEMENTATION PRIORITY MATRIX',
          heading: HeadingLevel.HEADING_1,
          spacing: { before: 100, after: 100 },
        }),
        ...createTable(
          ['Priority', 'Action', 'Keywords Impacted', 'Effort', 'Impact'],
          [
            ['P1', 'Fix H1 visibility (remove sr-only class)', 'All core keywords', 'Low', 'High'],
            ['P1', 'Add Product schema markup to product pages', 'All product keywords', 'Medium', 'Critical'],
            ['P1', 'Add BreadcrumbList schema markup', 'Category/product pages', 'Low', 'High'],
            ['P1', 'Create SEO-friendly category slugs', 'Category keywords', 'High', 'Critical'],
            ['P1', 'Add product pages to sitemap.xml', 'Product keywords', 'Medium', 'Critical'],
            ['P2', 'Rewrite homepage content (500+ words)', 'Core sourcing keywords', 'Medium', 'High'],
            ['P2', 'Add FAQPage schema markup', 'Import process keywords', 'Low', 'Medium'],
            ['P2', 'Add LocalBusiness schema + Google Business Profile', 'Local SEO keywords', 'Medium', 'High'],
            ['P3', 'Create blog content (10 articles)', 'Informational long-tail keywords', 'High', 'Medium'],
            ['P3', 'Bengali page translations + hreflang tags', 'Bengali keywords', 'Medium', 'Medium'],
            ['P3', 'Optimize meta descriptions for all pages', 'All pages', 'Low', 'Medium'],
            ['P4', 'Link building campaign', 'Authority keywords', 'High', 'Medium'],
          ]
        ),

        new Paragraph({
          text: '8. SUCCESS METRICS (KPIs)',
          heading: HeadingLevel.HEADING_1,
          spacing: { before: 200, after: 100 },
        }),
        ...createTable(
          ['Metric', 'Current', 'Target (6 Months)', 'Target (12 Months)'],
          [
            ['Organic traffic (monthly)', 'Unknown', '5,000+', '15,000+'],
            ['Keywords ranking top 10', '~5', '50+', '150+'],
            ['Product pages indexed', '0', '100+', '500+'],
            ['Category pages indexed', '~10', '30+', '50+'],
            ['Domain Rating', '~5-10', '20+', '35+'],
            ['Backlinks (referring domains)', '<50', '200+', '500+'],
            ['Bengali keyword rankings', '0', '20+', '50+'],
          ]
        ),
      ],
    }],
  });

  Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync('/workspace/app/WholesalePlus_Keyword_Research.docx', buffer);
    console.log('DOCX created successfully');
  });
}

function createTable(headers, rows) {
  const tableRows = [
    new TableRow({
      children: headers.map(h =>
        new TableCell({
          children: [new Paragraph({ text: h, bold: true, shading: { fill: '1F3864' }, color: 'FFFFFF' })],
          width: { size: 2000 / headers.length, type: WidthType.PERCENTAGE },
        })
      ),
    }),
    ...rows.map(row =>
      new TableRow({
        children: row.map(cell =>
          new TableCell({
            children: [new Paragraph({ text: String(cell) })],
            width: { size: 2000 / headers.length, type: WidthType.PERCENTAGE },
          })
        ),
      })
    ),
  ];

  return [
    new Table({
      rows: tableRows,
      width: { size: 100, type: WidthType.PERCENTAGE },
    }),
    new Paragraph({ text: '', spacing: { after: 100 } }),
  ];
}

createXLSX();
createDOCX();
