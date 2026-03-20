let mongoose = require('mongoose');

let inventorySchema = mongoose.Schema(
    {
        product: {
            type: mongoose.Types.ObjectId,
            ref: 'product',
            required: true,
            unique: true
        },
        stock: {
            type: Number,
            default: 0,
            min: [0, 'Stock cannot be negative']
        },
        reserved: {
            type: Number,
            default: 0,
            min: [0, 'Reserved cannot be negative']
        },
        soldCount: {
            type: Number,
            default: 0,
            min: [0, 'Sold count cannot be negative']
        }
    },
    {
        timestamps: true
    }
);

module.exports = mongoose.model('inventory', inventorySchema);
