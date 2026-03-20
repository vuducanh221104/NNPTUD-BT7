var express = require('express');
var router = express.Router();
let inventorySchema = require('../schemas/inventory');

// GET /api/v1/inventory - get all inventories with product join
router.get('/', async function (req, res, next) {
    try {
        let data = await inventorySchema.find({}).populate({
            path: 'product',
            select: 'title slug price description images category'
        });
        let result = data.filter(function (e) {
            return e.product && !e.product.isDeleted;
        });
        res.send(result);
    } catch (error) {
        res.status(500).send({ message: error.message });
    }
});

// GET /api/v1/inventory/:id - get inventory by ID with product join
router.get('/:id', async function (req, res, next) {
    try {
        let result = await inventorySchema.findOne({ _id: req.params.id }).populate({
            path: 'product',
            select: 'title slug price description images category'
        });
        if (result) {
            res.status(200).send(result);
        } else {
            res.status(404).send({ message: 'INVENTORY NOT FOUND' });
        }
    } catch (error) {
        res.status(404).send({ message: 'INVENTORY NOT FOUND' });
    }
});

// POST /api/v1/inventory/add_stock - add stock to a product
router.post('/add_stock', async function (req, res, next) {
    try {
        let { product, quantity } = req.body;
        if (!product || quantity === undefined || quantity < 0) {
            return res.status(400).send({ message: 'Invalid product or quantity' });
        }
        let inventory = await inventorySchema.findOne({ product });
        if (!inventory) {
            return res.status(404).send({ message: 'Inventory not found for this product' });
        }
        inventory.stock += Number(quantity);
        await inventory.save();
        let updated = await inventorySchema.findOne({ product }).populate({
            path: 'product',
            select: 'title slug price description images category'
        });
        res.status(200).send(updated);
    } catch (error) {
        res.status(500).send({ message: error.message });
    }
});

// POST /api/v1/inventory/remove_stock - remove stock from a product
router.post('/remove_stock', async function (req, res, next) {
    try {
        let { product, quantity } = req.body;
        if (!product || quantity === undefined || quantity < 0) {
            return res.status(400).send({ message: 'Invalid product or quantity' });
        }
        let inventory = await inventorySchema.findOne({ product });
        if (!inventory) {
            return res.status(404).send({ message: 'Inventory not found for this product' });
        }
        if (inventory.stock < Number(quantity)) {
            return res.status(400).send({ message: 'Insufficient stock' });
        }
        inventory.stock -= Number(quantity);
        await inventory.save();
        let updated = await inventorySchema.findOne({ product }).populate({
            path: 'product',
            select: 'title slug price description images category'
        });
        res.status(200).send(updated);
    } catch (error) {
        res.status(500).send({ message: error.message });
    }
});

// POST /api/v1/inventory/reservation - reserve stock (decrease stock, increase reserved)
router.post('/reservation', async function (req, res, next) {
    try {
        let { product, quantity } = req.body;
        if (!product || quantity === undefined || quantity < 0) {
            return res.status(400).send({ message: 'Invalid product or quantity' });
        }
        let inventory = await inventorySchema.findOne({ product });
        if (!inventory) {
            return res.status(404).send({ message: 'Inventory not found for this product' });
        }
        if (inventory.stock < Number(quantity)) {
            return res.status(400).send({ message: 'Insufficient stock to reserve' });
        }
        inventory.stock -= Number(quantity);
        inventory.reserved += Number(quantity);
        await inventory.save();
        let updated = await inventorySchema.findOne({ product }).populate({
            path: 'product',
            select: 'title slug price description images category'
        });
        res.status(200).send(updated);
    } catch (error) {
        res.status(500).send({ message: error.message });
    }
});

// POST /api/v1/inventory/sold - mark as sold (decrease reserved, increase soldCount)
router.post('/sold', async function (req, res, next) {
    try {
        let { product, quantity } = req.body;
        if (!product || quantity === undefined || quantity < 0) {
            return res.status(400).send({ message: 'Invalid product or quantity' });
        }
        let inventory = await inventorySchema.findOne({ product });
        if (!inventory) {
            return res.status(404).send({ message: 'Inventory not found for this product' });
        }
        if (inventory.reserved < Number(quantity)) {
            return res.status(400).send({ message: 'Insufficient reserved quantity to sell' });
        }
        inventory.reserved -= Number(quantity);
        inventory.soldCount += Number(quantity);
        await inventory.save();
        let updated = await inventorySchema.findOne({ product }).populate({
            path: 'product',
            select: 'title slug price description images category'
        });
        res.status(200).send(updated);
    } catch (error) {
        res.status(500).send({ message: error.message });
    }
});

module.exports = router;
